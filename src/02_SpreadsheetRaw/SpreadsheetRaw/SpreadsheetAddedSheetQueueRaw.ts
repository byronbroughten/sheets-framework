import { Arr } from "@byronbroughten/utils/arr";
import { Val } from "@byronbroughten/utils/val";

import type {
  AddSheetOperation,
  BoundedGridRange,
  GridRangeProps,
} from "../../00_Source/RawSource/RawSource.ts";
import { headRows } from "../../01_SpreadsheetSchema/headRows.ts";
import { SpreadsheetBaseRaw } from "../ClassBases/SpreadsheetBaseRaw.ts";
import type { AddedSheetCell, GridEnd } from "../ClassTypes/StateRaw.ts";

export class SpreadsheetAddedSheetQueueRaw extends SpreadsheetBaseRaw {
  // A tab this flush adds has no sheet state until fetched, so its queued addSheet is what puts it in the working view.
  sheetIsAdded(sheetGid: number): boolean {
    return this._queuedAddSheet(sheetGid) !== undefined;
  }
  // So a further Table this flush adds on the tab fits its grid.
  updateAddedSheetGridThrough(
    sheetGid: number,
    { endRowIndex, endColumnIndex }: GridEnd,
  ): this {
    const addSheet = Val.assert(
      this._queuedAddSheet(sheetGid),
      `queued addSheet for GID ${sheetGid}`,
    );
    addSheet.rowCount = Math.max(addSheet.rowCount, endRowIndex);
    addSheet.columnCount = Math.max(addSheet.columnCount, endColumnIndex);
    return this;
  }
  addedTablesEndColIndex(sheetGid: number): number | undefined {
    return Arr.max(
      this._queuedAddTableRanges(sheetGid).map(
        ({ endColumnIndex }) => endColumnIndex,
      ),
    );
  }
  validateAddSheetQueued(sheetId: number, write: string): void {
    if (this.sheetIsAdded(sheetId)) return;
    throw new Error(
      `Added-sheet ${write} refused: no addSheet for GID ${sheetId} is queued in this flush.`,
    );
  }
  validateInAddedTable({ sheetId, rowIndex, colIndex }: AddedSheetCell): void {
    const added = this._addedTableHolds(
      sheetId,
      (range) =>
        range.startRowIndex - headRows.countAboveHeader <= rowIndex &&
        rowIndex < range.endRowIndex &&
        range.startColumnIndex <= colIndex &&
        colIndex < range.endColumnIndex,
    );
    if (added) return;
    throw new Error(
      `Added-Table cell write refused: no addTable queued in this flush holds row ${rowIndex}, column ${colIndex} of GID ${sheetId}.`,
    );
  }
  validateRangeInAddedTable(rangeProps: GridRangeProps): void {
    const {
      sheetId,
      startRowIndex,
      endRowIndex,
      startColumnIndex,
      endColumnIndex,
    } = rangeProps;
    const added = this._addedTableHolds(
      sheetId,
      (range) =>
        range.startRowIndex <= startRowIndex &&
        endRowIndex !== undefined &&
        endRowIndex <= range.endRowIndex &&
        startColumnIndex !== undefined &&
        range.startColumnIndex <= startColumnIndex &&
        endColumnIndex !== undefined &&
        endColumnIndex <= range.endColumnIndex,
    );
    if (added) return;
    throw new Error(
      `Added-Table rule refused: no addTable queued in this flush holds the range ${JSON.stringify(rangeProps)}.`,
    );
  }
  private _queuedAddSheet(sheetGid: number): AddSheetOperation | undefined {
    return this.writeOperations.addSheet.find(
      ({ sheetId }) => sheetId === sheetGid,
    );
  }
  private _addedTableHolds(
    sheetId: number,
    holds: (range: BoundedGridRange) => boolean,
  ): boolean {
    return this._queuedAddTableRanges(sheetId).some(holds);
  }
  private _queuedAddTableRanges(sheetGid: number): BoundedGridRange[] {
    return this.writeOperations.addTable
      .map(({ range }) => range)
      .filter(({ sheetId }) => sheetId === sheetGid);
  }
}
