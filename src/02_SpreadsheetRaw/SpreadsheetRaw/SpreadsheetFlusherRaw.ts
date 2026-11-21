import { PartialTableRefusal } from "../../00_Source/RawSource/PartialTableRefusal";
import type {
  AppendDimensionOperation,
  DeleteConditionalFormatRuleOperation,
  DeleteTableRowsOperation,
  LocalWriteOperation,
} from "../../00_Source/RawSource/RawSource";
import { SpreadsheetBaseRaw } from "../ClassBases/SpreadsheetBaseRaw";
import { emptyStateRaw } from "../ClassTypes/emptyStateRaw";
import type {
  AppendTableRows,
  RowWrites,
  SheetWriteQueueRaw,
  TableColumnChangeOperation,
  TableColumnInsertOperation,
  TableGrowthOperation,
  TableStateRaw,
} from "../ClassTypes/StateRaw";
import { SpreadsheetRaw } from "../SpreadsheetRaw";
import type { TableRaw } from "../TableRaw";
import { SpreadsheetRefusalRewriterRaw } from "./SpreadsheetRefusalRewriterRaw";

interface QueuedRowWrites {
  rowIndex: number;
  writes: RowWrites;
}

export class SpreadsheetFlusherRaw extends SpreadsheetBaseRaw {
  get ss(): SpreadsheetRaw {
    return new SpreadsheetRaw(this.spreadsheetRawProps);
  }
  private get refusalRewriter(): SpreadsheetRefusalRewriterRaw {
    return new SpreadsheetRefusalRewriterRaw(this.spreadsheetRawProps);
  }
  flush(): void {
    // Before the gather, which empties the Table queues it reads.
    const tableIdsWithColumnTypeUpdates = this._tableIdsWhere(
      ({ writeQueue }) => writeQueue.table.columnTypes.size > 0,
    );
    const tableIdsWithSorts = this._tableIdsWhere(
      ({ writeQueue }) => writeQueue.table.sort !== undefined,
    );
    // Before the gather too, whose shifts would misplace a refused run's failure status.
    for (const tableId of tableIdsWithColumnTypeUpdates) {
      this.ss.table(tableId).assertDropdownStyleResetAllowed();
    }
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
    sheetGidsWithRowDeletes.forEach((sheetGid) =>
      this.ss.sheet(sheetGid).markRowIndexesStale(),
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
  private _tableIdsWhere(
    isPicked: (state: TableStateRaw) => boolean,
  ): string[] {
    return Array.from(this.tablesStateRaw.entries())
      .filter(([, state]) => isPicked(state))
      .map(([tableId]) => tableId);
  }
  private _gatherWriteOperations(): void {
    const tables = this._tablesWithWriteQueues();
    tables.forEach((table) => table.gatherAppendTableRowsOperation());
    this._shiftTablesBelowGrowth(this.writeOperations.appendTableRows);
    tables.forEach((table) => table.gatherInsertTableEndColumnsOperation());
    const { insertTableEndColumns, changeTableColumns } = this.writeOperations;
    this._shiftTablesRight((table) =>
      table.columnShiftFrom(insertTableEndColumns),
    );
    for (const table of tables) table.gatherColumnChangesOperation();
    // Negative where a Table's deletes outnumber its inserts, which shifts the Tables to its right left.
    this._shiftTablesRight((table) =>
      table.columnShiftFromChanges(changeTableColumns),
    );
    tables.forEach((table) => {
      table.gatherQueuedTableWrites();
      for (const [rowIndex, writes] of table.rowWrites) {
        gatherRowWrites(table, { rowIndex, writes });
      }
    });
    this._shiftTablesBelowRowDeletes(this.writeOperations.deleteTableRows);
    tables.forEach((table) => {
      // After the delete shift, since sorts are the one sheet-coordinate write sent after the deletes.
      table.gatherSortTableOperation();
      // After the Table queues, so the insert-column refusal sees this flush's inserts.
      table.gatherSetTableColumnPropertiesOperation();
      table._clearWriteQueue();
    });
  }
  private _shiftTablesBelowGrowth(growths: AppendTableRows[]): void {
    this._measuredShifts((table) => table.rowShiftFrom(growths)).forEach(
      ({ table, shiftCount }) => table.shiftRowsDown(shiftCount),
    );
  }
  private _shiftTablesBelowRowDeletes(
    deletes: DeleteTableRowsOperation[],
  ): void {
    this._measuredShifts((table) => table.rowPullFrom(deletes)).forEach(
      ({ table, shiftCount }) => table.shiftRowsUp(shiftCount),
    );
  }
  private _shiftTablesRight(shiftOf: (table: TableRaw) => number): void {
    for (const { table, shiftCount } of this._measuredShifts(shiftOf)) {
      table.shiftColumnsRight(shiftCount);
    }
  }
  // Measured against the layout before the batch's shifting operations, then applied, so no shift sees another.
  private _measuredShifts(
    shiftOf: (table: TableRaw) => number,
  ): { table: TableRaw; shiftCount: number }[] {
    return Array.from(this.tablesStateRaw.keys(), (tableId) => {
      const table = this.ss.table(tableId);
      return { table, shiftCount: shiftOf(table) };
    }).filter(({ shiftCount }) => shiftCount !== 0);
  }
  // Listed so a queue on a Table not yet fetched stops the run at gathering instead of vanishing.
  private _tablesWithWriteQueues(): TableRaw[] {
    const knownTables = Array.from(this.tablesStateRaw.keys(), (tableId) =>
      this.ss.table(tableId),
    );
    const tablesBeforePropertiesById = Array.from(
      this.sheetsStateRaw.entries(),
    ).flatMap(([sheetGid, sheetState]) =>
      Array.from(sheetState.tablesBeforePropertiesById.keys(), (tableId) =>
        this.ss.sheet(sheetGid).table(tableId),
      ),
    );
    return [...knownTables, ...tablesBeforePropertiesById];
  }
  private _sheetGidsWithRowDeletes(): Set<number> {
    return new Set(
      this.writeOperations.deleteTableRows.map(({ range }) => range.sheetId),
    );
  }
  // A column delete removes a rule or protection it holds whole, so the sheet's fetched ones no longer match.
  private _sheetGidsWithColumnChanges(): number[] {
    return this.writeOperations.changeTableColumns.map(
      ({ sheetId }) => sheetId,
    );
  }
  private _sheetGidsWithConditionalFormatMutations(): Set<number> {
    const sheetGids = new Set<number>(this._sheetGidsWithColumnChanges());
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
    // A tab this flush adds has no sheet state to mark; its first fetch reads its rules fresh.
    for (const { sheetId } of this.writeOperations.addSheet) {
      sheetGids.delete(sheetId);
    }
    return sheetGids;
  }
  private _sheetGidsWithEditProtectionMutations(): Set<number> {
    return new Set([
      ...this._sheetGidsWithColumnChanges(),
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
      // A new Table needs its tab and its grid in place first; added past the grid's edge, it fails.
      ...queued.addSheet,
      ...this._appendDimensionOperations(),
      ...queued.addTable,
      ...queued.renameSheet,
      ...queued.updateTabProperties,
      ...queued.renameTable,
      // First among column writes, so a header write in the same batch renames the column rather than being reverted.
      ...queued.setTableColumnProperties,
      ...this._appendTableRowsOperationsBottomUp(),
      ...this._insertTableEndColumnsOperationsRightToLeft(),
      // After the Table-end inserts, whose layout each change's band was gathered in.
      ...this._changeTableColumnsOperationsRightToLeft(),
      // Column fills go before cell writes; a later-queued fill already erased the cell writes it covers.
      ...queued.fillColumn,
      ...queued.fillCell,
      // After cell writes, so a checkbox's seeded value is written before its rule.
      ...queued.addCheckboxValidation,
      // Reads the text as it stands mid-batch, so it must follow what writes it.
      ...queued.findReplace,
      // Before the row deletes, which carry them up with the rows, since they were gathered against the layout before.
      ...this._deleteConditionalFormatOperationsDescending(),
      ...queued.addConditionalFormatRule,
      ...queued.updateConditionalFormatRule,
      ...queued.deleteProtectedRange,
      ...queued.addProtectedRange,
      // Gathered against the layout before the row deletes, like the rules above.
      ...queued.deleteTable,
      ...this._deleteOperationsDescending(),
      ...queued.sortTable,
      // After every write, so none queued before it lands on a tab already gone.
      ...queued.deleteSheet,
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
      throw this.refusalRewriter.rewordedRefusal(error);
    }
    if (operations.length > 0) this.spreadsheetStateRaw.landedFlushCount += 1;
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
      .sort(
        bySheetThenPositionDescending(({ newRows }) => ({
          sheetId: newRows.sheetId,
          position: newRows.startRowIndex,
        })),
      )
      .flatMap(({ operations }) => operations);
  }
  // Rightmost Table first within a sheet, so no insert shifts a Table whose insert is still to come.
  private _insertTableEndColumnsOperationsRightToLeft(): TableColumnInsertOperation[] {
    return operationsRightToLeft(
      this.writeOperations.insertTableEndColumns.map(
        ({ newColumns, operations }) => ({ ...newColumns, operations }),
      ),
    );
  }
  // Rightmost Table first within a sheet, so no change moves a Table whose own are still to come; each Table's in queue order.
  private _changeTableColumnsOperationsRightToLeft(): TableColumnChangeOperation[] {
    return operationsRightToLeft(this.writeOperations.changeTableColumns);
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
      bySheetThenPositionDescending(({ sheetId, index }) => ({
        sheetId,
        position: index,
      })),
    );
  }
  // Scope can be allSheets, so one rule: every Table's fetched cells go stale.
  private _invalidateFetchedCellState(): void {
    this.tablesStateRaw.forEach((_, tableId) =>
      this.ss.table(tableId).invalidateCellState(),
    );
  }
}

function gatherRowWrites(
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

interface SheetPlace {
  sheetId: number;
  position: number;
}

function bySheetThenPositionDescending<OP extends object>(
  placeOf: (operation: OP) => SheetPlace,
): (a: OP, b: OP) => number {
  return (a, b) => {
    const placeA = placeOf(a);
    const placeB = placeOf(b);
    if (placeA.sheetId !== placeB.sheetId) {
      return placeA.sheetId - placeB.sheetId;
    }
    return placeB.position - placeA.position;
  };
}

function operationsRightToLeft<OP>(
  columnWrites: {
    sheetId: number;
    startColumnIndex: number;
    operations: OP[];
  }[],
): OP[] {
  return [...columnWrites]
    .sort(
      bySheetThenPositionDescending(({ sheetId, startColumnIndex }) => ({
        sheetId,
        position: startColumnIndex,
      })),
    )
    .flatMap(({ operations }) => operations);
}
