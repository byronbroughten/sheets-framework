import type {
  DeleteConditionalFormatRuleOperation,
  DeleteRowsOperation,
} from "../../00_Source/RawSource/RawSource";
import { SpreadsheetBaseRaw } from "../ClassBases/SpreadsheetBaseRaw";
import { emptyStateRaw } from "../ClassTypes/emptyStateRaw";
import type { RowWrites } from "../ClassTypes/StateRaw";
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
    // Row indexes only actually shift once the deletes have been sent, and a sheet-row delete shifts every Table on the sheet.
    sheetGidsWithRowDeletes.forEach((sheetGid) =>
      this.ss
        .sheet(sheetGid)
        .tableIds()
        .forEach((tableId) => this.ss.table(tableId).markRowIndexesStale()),
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
  private _gatherWriteOperations(): void {
    const tables = this._tablesWithWriteQueues();
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
  // A Table not yet fetched holds its queue on its sheet, aimed where the layout expects it.
  private _tablesWithWriteQueues(): TableRaw[] {
    const knownTables = Array.from(this.tablesStateRaw.keys(), (tableId) =>
      this.ss.table(tableId),
    );
    const tablesBeforeProperties = Array.from(
      this.sheetsStateRaw.keys(),
      (sheetGid) => this.ss.sheet(sheetGid),
    ).filter((sheet) => !sheet.hasFetchedProperties);
    return [...knownTables, ...tablesBeforeProperties];
  }
  private _gatherRowWrites(
    table: TableRaw,
    { rowIndex, writes }: QueuedRowWrites,
  ): void {
    if (writes.appendRow && writes.deleteRow) {
      return;
    } else if (writes.deleteRow) {
      const origin = table.tableOrigin();
      this.writeOperations.deleteRows.push({
        kind: "deleteRows",
        sheetId: table.sheetGid,
        startIndex: origin.sheetRowIndex(rowIndex),
        endIndex: origin.sheetRowIndex(rowIndex + 1),
      });
    } else {
      const row = table.rowCommon(rowIndex);
      if (writes.appendRow) {
        row.gatherAppendRowsOperation();
      }
      for (const [colIndex, cellFill] of writes.fillCells) {
        row.cell(colIndex).gatherFillCellOperation(cellFill);
      }
    }
  }
  private _sheetGidsWithRowDeletes(): Set<number> {
    return new Set(
      this.writeOperations.deleteRows.map(({ sheetId }) => sheetId),
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
      ...queued.appendRows,
      ...queued.insertTableEndColumn,
      // Column fills go before cell writes; a later-queued fill already erased the cell writes it covers.
      ...queued.fillColumn,
      ...queued.fillCell,
      // After cell writes, so a checkbox's seeded value is written before its rule.
      ...queued.addCheckboxValidation,
      // Reads the text as it stands mid-batch, so it must follow what writes it.
      ...queued.findReplace,
      ...this._deleteOperationsDescending(),
      ...queued.sort,
      ...this._deleteConditionalFormatOperationsDescending(),
      ...queued.addConditionalFormatRule,
      ...queued.deleteProtectedRange,
      ...queued.addProtectedRange,
      // Outside the ordering rules the queue was built around, so last.
      ...queued.raw,
    ];
    this.spreadsheetStateRaw.rawSource.flush(operations);
    this.spreadsheetStateRaw.writeQueue.operations =
      emptyStateRaw.writeOperations();
  }
  // Deletes within one batchUpdate apply sequentially and each shifts the
  // row indices below it, so same-sheet deletes must go highest-index-first
  // or a later request's pre-computed startIndex lands on the wrong row.
  private _deleteOperationsDescending(): DeleteRowsOperation[] {
    return [...this.writeOperations.deleteRows].sort(
      (a, b) => b.startIndex - a.startIndex,
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
