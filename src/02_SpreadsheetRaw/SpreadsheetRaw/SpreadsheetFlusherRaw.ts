import type {
  DeleteConditionalFormatRuleOperation,
  DeleteRowsOperation,
} from "../../00_Source/RawSource/RawSource";
import { SpreadsheetBaseRaw } from "../ClassBases/SpreadsheetBaseRaw";
import { emptyStateRaw } from "../ClassTypes/emptyStateRaw";
import type {
  RowChangesToSave,
  SheetChangesToSave,
  UpdateTableColumnTypeOperation,
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
    this._gatherUpdateRequests();
    const sheetGidsWithRowDeletes = this._sheetGidsWithRowDeletes();
    const sheetGidsWithConditionalFormatMutations =
      this._sheetGidsWithConditionalFormatMutations();
    const sheetGidsWithEditProtectionMutations =
      this._sheetGidsWithEditProtectionMutations();
    const hasFindReplace = this.updateRequests.findReplace.length > 0;
    const sheetGidsWithColumnTypeUpdates = new Set(
      this.updateRequests.updateTableColumnType.map(({ sheetId }) => sheetId),
    );
    this._sendUpdateRequests();
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
  private _gatherUpdateRequests(): void {
    this.sheetsStateRaw.forEach((state, sheetGid) => {
      this._gatherSheetRequests(sheetGid, state.writeQueue.sheet);
      for (const [rowIndex, change] of state.writeQueue.rows) {
        this._gatherRowRequests(change, { sheetGid, rowIndex });
      }
      state.writeQueue.sheet = emptyStateRaw.sheetChanges();
      state.writeQueue.rows = new Map();
    });
    // After the sheet queues, so the insert-column refusal sees this flush's inserts.
    this._gatherColumnTypesRequests();
  }
  private _gatherColumnTypesRequests(): void {
    const opsBySheet = new Map<number, UpdateTableColumnTypeOperation[]>();
    this.updateRequests.updateTableColumnType.forEach((operation) => {
      const ops = opsBySheet.get(operation.sheetId) ?? [];
      ops.push(operation);
      opsBySheet.set(operation.sheetId, ops);
    });
    opsBySheet.forEach((ops, sheetGid) =>
      this.ss.sheet(sheetGid).gatherColumnTypesRequest(ops),
    );
  }
  private _gatherSheetRequests(
    sheetGid: number,
    change: SheetChangesToSave,
  ): void {
    change.insertColumn.forEach(({ startColumnIndex }) => {
      this.ss.sheet(sheetGid).gatherInsertColumnRequest(startColumnIndex);
    });
    if (change.sort !== undefined) {
      this.ss.sheet(sheetGid).gatherSortRequest(change.sort);
    }
    change.fills.forEach((fill) => {
      this.ss.sheet(sheetGid).gatherFillRequest(fill);
    });
  }
  private _gatherRowRequests(
    change: RowChangesToSave,
    { sheetGid, rowIndex }: SheetRowRef,
  ): void {
    if (change.append && change.delete) {
      return;
    } else if (change.delete) {
      this.updateRequests.delete.push({
        kind: "deleteRows",
        sheetId: sheetGid,
        startIndex: rowIndex,
        endIndex: rowIndex + 1,
      });
    } else {
      const row = this.ss.sheet(sheetGid).rowCommon(rowIndex);
      if (change.append) {
        row.gatherAppendRequest();
      }
      for (const [colIndex, cellChange] of change.update) {
        row.cell(colIndex).gatherUpdateRequest(cellChange);
      }
    }
  }
  private _sheetGidsWithRowDeletes(): Set<number> {
    return new Set(this.updateRequests.delete.map(({ sheetId }) => sheetId));
  }
  private _sheetGidsWithConditionalFormatMutations(): Set<number> {
    const sheetGids = new Set<number>();
    this.updateRequests.deleteConditionalFormat.forEach(({ sheetId }) =>
      sheetGids.add(sheetId),
    );
    this.updateRequests.addConditionalFormat.forEach(({ rule }) => {
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
      ...this.updateRequests.deleteProtectedRange.map(({ sheetId }) => sheetId),
      ...this.updateRequests.addProtectedRange.map(
        ({ protection }) => protection.range.sheetId,
      ),
    ]);
  }
  private _sendUpdateRequests(): void {
    const requests = this.updateRequests;
    const operations = [
      // A Table can only be added to a tab this batch has created, so both go first.
      ...requests.addSheet,
      ...requests.addTable,
      ...requests.updateSheetTitle,
      ...requests.updateTableName,
      // First among column writes, so a header write in the same batch renames the column rather than being reverted.
      ...requests.updateTableColumnProperties,
      ...requests.append,
      ...requests.insertColumn,
      // Fills go before updates, so a per-cell write on a filled column wins.
      ...requests.fill,
      ...requests.update,
      // After cell updates, so a checkbox's seeded value is written before its rule.
      ...requests.addCheckboxValidation,
      // Reads the text as it stands mid-batch, so it must follow what writes it.
      ...requests.findReplace,
      ...this._deleteOperationsDescending(),
      ...requests.sort,
      ...this._deleteConditionalFormatOperationsDescending(),
      ...requests.addConditionalFormat,
      ...requests.deleteProtectedRange,
      ...requests.addProtectedRange,
      // Outside the ordering rules the queue was built around, so last.
      ...requests.raw,
    ];
    this.spreadsheetStateRaw.rawSource.flush(operations);
    this.spreadsheetStateRaw.writeQueue.updateRequests =
      emptyStateRaw.updateRequests();
  }
  // Deletes within one batchUpdate apply sequentially and each shifts the
  // row indices below it, so same-sheet deletes must go highest-index-first
  // or a later request's pre-computed startIndex lands on the wrong row.
  private _deleteOperationsDescending(): DeleteRowsOperation[] {
    return [...this.updateRequests.delete].sort(
      (a, b) => b.startIndex - a.startIndex,
    );
  }
  private _deleteConditionalFormatOperationsDescending(): DeleteConditionalFormatRuleOperation[] {
    return [...this.updateRequests.deleteConditionalFormat].sort((a, b) => {
      if (a.sheetId !== b.sheetId) return a.sheetId - b.sheetId;
      return b.index - a.index;
    });
  }
  // Scope can be allSheets, so one rule: every sheet's fetched cells go stale.
  private _invalidateFetchedCellState(): void {
    this.sheetsStateRaw.forEach((_, sheetGid) =>
      this.ss.sheet(sheetGid).invalidateCellState(),
    );
  }
}
