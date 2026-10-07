import { PartialTableRefusal } from "../../00_Source/RawSource/PartialTableRefusal";
import type {
  AppendDimensionOperation,
  DeleteConditionalFormatRuleOperation,
  DeleteTableRowsOperation,
  InsertRangeOperation,
  LocalWriteOperation,
  TableSnapshot,
} from "../../00_Source/RawSource/RawSource";
import { SheetIndex } from "../../00_Source/RawSource/SheetIndex";
import { SpreadsheetBaseRaw } from "../ClassBases/SpreadsheetBaseRaw";
import { rowShiftFrom } from "../ClassBases/TableCommonRaw";
import { emptyStateRaw } from "../ClassTypes/emptyStateRaw";
import type {
  AppendTableRows,
  InsertTableEndColumns,
  RowWrites,
  SheetWriteQueueRaw,
  TableColumnInsertOperation,
  TableGrowthOperation,
} from "../ClassTypes/StateRaw";
import { SpreadsheetRaw } from "../SpreadsheetRaw";
import type { TableRaw } from "../TableRaw";

interface QueuedRowWrites {
  rowIndex: number;
  writes: RowWrites;
}

export class SpreadsheetFlusherRaw extends SpreadsheetBaseRaw {
  get ss(): SpreadsheetRaw {
    return new SpreadsheetRaw(this.spreadsheetRawProps);
  }
  flush(): void {
    // Before the gather, which empties the Table queues it reads.
    const tableIdsWithColumnTypeUpdates = this._tableIdsWithColumnTypeUpdates();
    const tableIdsWithSorts = this._tableIdsWithSorts();
    this._gatherWriteOperations();
    const sheetGidsWithRowDeletes = this._sheetGidsWithRowDeletes();
    const sheetGidsWithConditionalFormatMutations =
      this._sheetGidsWithConditionalFormatMutations();
    const sheetGidsWithEditProtectionMutations =
      this._sheetGidsWithEditProtectionMutations();
    const hasFindReplace = this.writeOperations.findReplace.length > 0;
    this._sendWriteOperations();
    tableIdsWithColumnTypeUpdates.forEach((tableId) =>
      this.ss.table(tableId).markColumnPropertiesStale(),
    );
    // Row indexes only actually shift once the deletes have been sent, and a Table below the deleted rows shifts with them.
    sheetGidsWithRowDeletes.forEach((sheetGid) =>
      this.ss
        .tableOnSheet(sheetGid)
        .tableIds()
        .forEach((tableId) => this.ss.table(tableId).markRowIndexesStale()),
    );
    tableIdsWithSorts.forEach((tableId) =>
      this.ss.table(tableId).markRowIndexesStale(),
    );
    sheetGidsWithConditionalFormatMutations.forEach((sheetGid) =>
      this.ss.sheet(sheetGid).markConditionalFormatIndexesStale(),
    );
    sheetGidsWithEditProtectionMutations.forEach((sheetGid) =>
      this.ss.sheet(sheetGid).markEditProtectionsStale(),
    );
    if (hasFindReplace) this._invalidateFetchedCellState();
  }
  private _tableIdsWithColumnTypeUpdates(): string[] {
    return Array.from(this.tablesStateRaw.entries())
      .filter(([, state]) => state.writeQueue.table.columnTypes.size > 0)
      .map(([tableId]) => tableId);
  }
  private _tableIdsWithSorts(): string[] {
    return Array.from(this.tablesStateRaw.entries())
      .filter(([, state]) => state.writeQueue.table.sort !== undefined)
      .map(([tableId]) => tableId);
  }
  private _gatherWriteOperations(): void {
    const tables = this._tablesWithWriteQueues();
    tables.forEach((table) => table.gatherAppendTableRowsOperation());
    this._shiftTablesBelowGrowth(this.writeOperations.appendTableRows);
    tables.forEach((table) => table.gatherInsertTableEndColumnsOperation());
    this._shiftTablesRightOfColumnInserts(
      this.writeOperations.insertTableEndColumns,
    );
    tables.forEach((table) => {
      table.gatherQueuedTableWrites();
      for (const [rowIndex, writes] of table.rowWrites) {
        this._gatherRowWrites(table, { rowIndex, writes });
      }
    });
    // After the Table queues, so the insert-column refusal sees this flush's inserts.
    tables.forEach((table) => {
      table.gatherSetTableColumnPropertiesOperation();
      table._clearWriteQueue();
    });
  }
  private _shiftTablesBelowGrowth(growths: AppendTableRows[]): void {
    this._measuredShifts((table) => table.rowShiftFrom(growths)).forEach(
      ({ table, shiftCount }) => table.shiftRowsDown(shiftCount),
    );
  }
  private _shiftTablesRightOfColumnInserts(
    inserts: InsertTableEndColumns[],
  ): void {
    this._measuredShifts((table) => table.columnShiftFrom(inserts)).forEach(
      ({ table, shiftCount }) => table.shiftColumnsRight(shiftCount),
    );
  }
  // Measured against the layout before the batch's shifting operations, then applied, so no shift sees another.
  private _measuredShifts(
    shiftOf: (table: TableRaw) => number,
  ): { table: TableRaw; shiftCount: number }[] {
    return Array.from(this.tablesStateRaw.keys(), (tableId) => {
      const table = this.ss.table(tableId);
      return { table, shiftCount: shiftOf(table) };
    }).filter(({ shiftCount }) => shiftCount > 0);
  }
  // A Table not yet fetched holds its queue on its sheet, aimed where the layout expects it.
  private _tablesWithWriteQueues(): TableRaw[] {
    const knownTables = Array.from(this.tablesStateRaw.keys(), (tableId) =>
      this.ss.table(tableId),
    );
    const tablesBeforeProperties = Array.from(
      this.sheetsStateRaw.keys(),
      (sheetGid) => this.ss.tableOnSheet(sheetGid),
    ).filter((table) => !table.hasFetchedProperties);
    return [...knownTables, ...tablesBeforeProperties];
  }
  private _gatherRowWrites(
    table: TableRaw,
    { rowIndex, writes }: QueuedRowWrites,
  ): void {
    const row = table.rowCommon(rowIndex);
    if (writes.deleteRow) {
      row.gatherDeleteTableRowsOperation();
    } else {
      for (const [colIndex, cellFill] of writes.fillCells) {
        row.cell(colIndex).gatherFillCellOperation(cellFill);
      }
    }
  }
  private _sheetGidsWithRowDeletes(): Set<number> {
    return new Set(
      this.writeOperations.deleteTableRows.map(({ range }) => range.sheetId),
    );
  }
  private _sheetGidsWithConditionalFormatMutations(): Set<number> {
    const sheetGids = new Set<number>();
    this.writeOperations.deleteConditionalFormatRule.forEach(({ sheetId }) =>
      sheetGids.add(sheetId),
    );
    this.writeOperations.addConditionalFormatRule.forEach(({ rule }) => {
      const sheetId = rule.ranges[0]?.sheetId;
      if (sheetId === undefined) {
        throw new Error(
          "Queued addConditionalFormatRule has no range sheetId.",
        );
      }
      sheetGids.add(sheetId);
    });
    return sheetGids;
  }
  private _sheetGidsWithEditProtectionMutations(): Set<number> {
    return new Set([
      ...this.writeOperations.deleteProtectedRange.map(
        ({ sheetId }) => sheetId,
      ),
      ...this.writeOperations.addProtectedRange.map(
        ({ protection }) => protection.range.sheetId,
      ),
    ]);
  }
  private _sendWriteOperations(): void {
    const queued = this.writeOperations;
    const operations = [
      // A Table can only be added to a tab this batch has created, so both go first.
      ...queued.addSheet,
      ...queued.addTable,
      ...queued.renameSheet,
      ...queued.renameTable,
      // First among column writes, so a header write in the same batch renames the column rather than being reverted.
      ...queued.setTableColumnProperties,
      ...this._appendDimensionOperations(),
      ...this._appendTableRowsOperationsBottomUp(),
      ...this._insertTableEndColumnsOperationsRightToLeft(),
      // Column fills go before cell writes; a later-queued fill already erased the cell writes it covers.
      ...queued.fillColumn,
      ...queued.fillCell,
      // After cell writes, so a checkbox's seeded value is written before its rule.
      ...queued.addCheckboxValidation,
      // Reads the text as it stands mid-batch, so it must follow what writes it.
      ...queued.findReplace,
      ...this._deleteOperationsDescending(),
      ...queued.sortTable,
      ...this._deleteConditionalFormatOperationsDescending(),
      ...queued.addConditionalFormatRule,
      ...queued.deleteProtectedRange,
      ...queued.addProtectedRange,
      // Outside the ordering rules the queue was built around, so last.
      ...queued.raw,
    ];
    this._flush(operations);
    this.spreadsheetStateRaw.writeQueue.operations =
      emptyStateRaw.writeOperations();
    this.sheetsStateRaw.forEach((state) => {
      state.writeQueue = emptyStateRaw.sheetWriteQueue();
    });
  }
  private _flush(operations: LocalWriteOperation[]): void {
    try {
      this.spreadsheetStateRaw.rawSource.flush(operations);
    } catch (error) {
      if (!(error instanceof PartialTableRefusal)) throw error;
      throw this._rewordedRefusal(error);
    }
  }
  // Google names neither Table, so the Table refused is the one the operation was gathered for.
  private _rewordedRefusal(refusal: PartialTableRefusal): Error {
    const { operation } = refusal;
    if (operation.kind !== "insertRange") return refusal;
    const split = this._fetchTableSplitBy(operation);
    if (split === undefined) return refusal;
    const growth = this.writeOperations.appendTableRows.find(({ operations }) =>
      operations.includes(operation),
    );
    if (growth !== undefined) {
      const grown = this.ss.table(growth.tableId);
      return new Error(
        `Growing Table "${grown.name}" on sheet "${grown.sheetTitle}" would insert cells over part of Table "${split.name}" below it. Make the lower Table, "${split.name}", no wider than "${grown.name}", or move it.`,
      );
    }
    const insert = this.writeOperations.insertTableEndColumns.find(
      ({ operations }) => operations.includes(operation),
    );
    if (insert === undefined) return refusal;
    const { startRowIndex, endRowIndex } = operation.range;
    return new Error(
      `${this.ss.table(insert.tableId).columnInsertSplitting(split.name)}. Move "${split.name}" so that it sits entirely within rows ${startRowIndex + 1}–${endRowIndex}, or entirely outside them.`,
    );
  }
  // Refused whole, so the live layout is from before this batch's growth, which a column insert is measured after.
  private _fetchTableSplitBy(
    insert: InsertRangeOperation,
  ): TableSnapshot | undefined {
    const { sheetId } = insert.range;
    const sheet = this.spreadsheetStateRaw.rawSource
      .fetchSheetProperties()
      .sheets.find(({ sheetGid }) => sheetGid === sheetId);
    const growths = this.writeOperations.appendTableRows;
    return sheet?.tables
      ?.map((table) => shiftedDownBy(growths, { ...table, sheetId }))
      .find((table) => isPartlyShiftedBy(insert, table));
  }
  // Sent before any insert, which can't start past the grid's edge.
  private _appendDimensionOperations(): AppendDimensionOperation[] {
    return Array.from(this.sheetsStateRaw).flatMap(
      ([sheetId, { writeQueue }]) =>
        gridAppends(sheetId, writeQueue).filter(
          ({ addedCount }) => addedCount > 0,
        ),
    );
  }
  // Bottom Table first within a sheet, so no insert shifts a Table whose growth is still to come.
  private _appendTableRowsOperationsBottomUp(): TableGrowthOperation[] {
    return [...this.writeOperations.appendTableRows]
      .sort(({ newRows: a }, { newRows: b }) => {
        if (a.sheetId !== b.sheetId) return a.sheetId - b.sheetId;
        return b.startRowIndex - a.startRowIndex;
      })
      .flatMap(({ operations }) => operations);
  }
  // Rightmost Table first within a sheet, so no insert shifts a Table whose insert is still to come.
  private _insertTableEndColumnsOperationsRightToLeft(): TableColumnInsertOperation[] {
    return [...this.writeOperations.insertTableEndColumns]
      .sort(({ newColumns: a }, { newColumns: b }) => {
        if (a.sheetId !== b.sheetId) return a.sheetId - b.sheetId;
        return b.startColumnIndex - a.startColumnIndex;
      })
      .flatMap(({ operations }) => operations);
  }
  // Deletes within one batchUpdate apply sequentially and each shifts the
  // row indices below it, so same-sheet deletes must go highest-index-first
  // or a later request's pre-computed startRowIndex lands on the wrong row.
  private _deleteOperationsDescending(): DeleteTableRowsOperation[] {
    return [...this.writeOperations.deleteTableRows].sort(
      (a, b) => b.range.startRowIndex - a.range.startRowIndex,
    );
  }
  private _deleteConditionalFormatOperationsDescending(): DeleteConditionalFormatRuleOperation[] {
    return [...this.writeOperations.deleteConditionalFormatRule].sort(
      (a, b) => {
        if (a.sheetId !== b.sheetId) return a.sheetId - b.sheetId;
        return b.index - a.index;
      },
    );
  }
  // Scope can be allSheets, so one rule: every Table's fetched cells go stale.
  private _invalidateFetchedCellState(): void {
    this.tablesStateRaw.forEach((_, tableId) =>
      this.ss.table(tableId).invalidateCellState(),
    );
  }
}

function shiftedDownBy(
  growths: AppendTableRows[],
  table: TableSnapshot & { sheetId: number },
): TableSnapshot {
  const rowCount = rowShiftFrom(growths, table);
  return {
    ...table,
    startRowIndex: SheetIndex.row(table.startRowIndex + rowCount),
    endRowIndex: SheetIndex.row(table.endRowIndex + rowCount),
  };
}

// Google's rule: an insert reaching a Table must span all of it across the shift.
function isPartlyShiftedBy(
  { range, shiftDimension }: InsertRangeOperation,
  table: TableSnapshot,
): boolean {
  if (shiftDimension === "ROWS") {
    return (
      table.endRowIndex > range.startRowIndex &&
      coversPartOf(
        [range.startColumnIndex, range.endColumnIndex],
        [table.startColumnIndex, table.endColumnIndex],
      )
    );
  }
  return (
    table.endColumnIndex > range.startColumnIndex &&
    coversPartOf(
      [range.startRowIndex, range.endRowIndex],
      [table.startRowIndex, table.endRowIndex],
    )
  );
}

function coversPartOf(
  [bandStart, bandEnd]: [number, number],
  [start, end]: [number, number],
): boolean {
  const overlaps = bandStart < end && start < bandEnd;
  const coversAll = bandStart <= start && end <= bandEnd;
  return overlaps && !coversAll;
}

function gridAppends(
  sheetId: number,
  writeQueue: SheetWriteQueueRaw,
): AppendDimensionOperation[] {
  return [
    {
      kind: "appendDimension",
      sheetId,
      dimension: "ROWS",
      addedCount: writeQueue.appendedRowCount,
    },
    {
      kind: "appendDimension",
      sheetId,
      dimension: "COLUMNS",
      addedCount: writeQueue.appendedColumnCount,
    },
  ];
}
