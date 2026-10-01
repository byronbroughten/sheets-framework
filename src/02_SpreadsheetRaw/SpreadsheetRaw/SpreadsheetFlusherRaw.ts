import type {
  DeleteConditionalFormatRuleOperation,
  DeleteRowsOperation,
} from "../../00_Source/RawSource/RawSource";
import { SpreadsheetBaseRaw } from "../ClassBases/SpreadsheetBaseRaw";
import { emptyStateRaw } from "../ClassTypes/emptyStateRaw";
import type {
  RowWrites,
  SetTableColumnTypeOperation,
  SheetWrites,
} from "../ClassTypes/StateRaw";
import { SpreadsheetRaw } from "../SpreadsheetRaw";

interface SheetRowRef {
  sheetGid: number;
  rowIndex: number;
}

export class SpreadsheetFlusherRaw extends SpreadsheetBaseRaw {
  get ss(): SpreadsheetRaw {
    return new SpreadsheetRaw(this.spreadsheetRawProps);
  }
  flush(): void {
    this._gatherWriteOperations();
    const sheetGidsWithRowDeletes = this._sheetGidsWithRowDeletes();
    const sheetGidsWithConditionalFormatMutations =
      this._sheetGidsWithConditionalFormatMutations();
    const sheetGidsWithEditProtectionMutations =
      this._sheetGidsWithEditProtectionMutations();
    const hasFindReplace = this.writeOperations.findReplace.length > 0;
    const sheetGidsWithColumnTypeUpdates = new Set(
      this.writeOperations.setTableColumnType.map(({ sheetId }) => sheetId),
    );
    this._sendWriteOperations();
    sheetGidsWithColumnTypeUpdates.forEach((sheetGid) =>
      this.ss.sheet(sheetGid).markColumnPropertiesStale(),
    );
    // Row indexes only actually shift once the deletes have been sent.
    sheetGidsWithRowDeletes.forEach((sheetGid) =>
      this.ss.sheet(sheetGid).markRowIndexesStale(),
    );
    sheetGidsWithConditionalFormatMutations.forEach((sheetGid) =>
      this.ss.sheet(sheetGid).markConditionalFormatIndexesStale(),
    );
    sheetGidsWithEditProtectionMutations.forEach((sheetGid) =>
      this.ss.sheet(sheetGid).markEditProtectionsStale(),
    );
    if (hasFindReplace) this._invalidateFetchedCellState();
  }
  private _gatherWriteOperations(): void {
    this.sheetsStateRaw.forEach((state, sheetGid) => {
      this._gatherSheetWrites(sheetGid, state.writeQueue.sheet);
      for (const [rowIndex, writes] of state.writeQueue.rows) {
        this._gatherRowWrites(writes, { sheetGid, rowIndex });
      }
      state.writeQueue.sheet = emptyStateRaw.sheetWrites();
      state.writeQueue.rows = new Map();
    });
    // After the sheet queues, so the insert-column refusal sees this flush's inserts.
    this._gatherSetTableColumnTypeOperations();
  }
  private _gatherSetTableColumnTypeOperations(): void {
    const opsBySheet = new Map<number, SetTableColumnTypeOperation[]>();
    this.writeOperations.setTableColumnType.forEach((operation) => {
      const ops = opsBySheet.get(operation.sheetId) ?? [];
      ops.push(operation);
      opsBySheet.set(operation.sheetId, ops);
    });
    opsBySheet.forEach((ops, sheetGid) =>
      this.ss.sheet(sheetGid).gatherSetTableColumnTypeOperation(ops),
    );
  }
  private _gatherSheetWrites(sheetGid: number, writes: SheetWrites): void {
    this.ss
      .sheet(sheetGid)
      .gatherInsertTableEndColumnOperations(writes.insertTableEndColumnCount);
    if (writes.sort !== undefined) {
      this.ss.sheet(sheetGid).gatherSortOperation(writes.sort);
    }
    writes.fillColumns.forEach((fill) => {
      this.ss.sheet(sheetGid).gatherFillColumnOperation(fill);
    });
  }
  private _gatherRowWrites(
    writes: RowWrites,
    { sheetGid, rowIndex }: SheetRowRef,
  ): void {
    if (writes.appendRow && writes.deleteRow) {
      return;
    } else if (writes.deleteRow) {
      this.writeOperations.deleteRows.push({
        kind: "deleteRows",
        sheetId: sheetGid,
        startIndex: rowIndex,
        endIndex: rowIndex + 1,
      });
    } else {
      const row = this.ss.sheet(sheetGid).rowCommon(rowIndex);
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
      // After cell updates, so a checkbox's seeded value is written before its rule.
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
  // Scope can be allSheets, so one rule: every sheet's fetched cells go stale.
  private _invalidateFetchedCellState(): void {
    this.sheetsStateRaw.forEach((_, sheetGid) =>
      this.ss.sheet(sheetGid).invalidateCellState(),
    );
  }
}
