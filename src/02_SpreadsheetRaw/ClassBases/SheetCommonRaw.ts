import { Arr } from "../../utils/Arr";
import { Obj } from "../../utils/Obj";
import { ActiveTableRaw } from "../ActiveTableRaw";
import type { SheetGridRangeProps } from "../ClassTypes/AccessorsRaw";
import type {
  SheetChangeProps,
  SheetChangesToSave,
} from "../ClassTypes/StateRaw";
import type { SpreadsheetRaw } from "../SpreadsheetRaw";
import { SheetBaseRaw } from "./SheetBaseRaw";

// Not on SheetBaseRaw: the row and column base classes hang off that.
export abstract class SheetCommonRaw extends SheetBaseRaw {
  // Abstract: importing SpreadsheetRaw here would close an init-time cycle.
  abstract get ss(): SpreadsheetRaw;
  get activeTable(): ActiveTableRaw {
    const knownTable = this.sheetState.working.knownTable;
    if (knownTable === undefined) {
      throw new Error(
        `Active table is null for sheetGid ${this.sheetGid}. Ensure that the sheet properties have been fetched.`,
      );
    }
    if (knownTable.endRowIndex <= this.schema.topDataRowIdx) {
      throw new Error(
        `Sheet ${this.sheetLabel} Table must have at least one data row.`,
      );
    }
    return new ActiveTableRaw(this.sheetRawProps);
  }
  get fullTableColIndexes(): number[] {
    return Arr.indexesFromUntil(
      this.activeTable.startColumnIndex,
      this.activeTable.endColumnIndex,
    );
  }
  get changesToSave(): SheetChangesToSave {
    return this.sheetState.writeQueue.sheet;
  }
  // The table's own range, not the layout's: no table means no table columns.
  isTableColIndex(colIndex: number): boolean {
    if (this.sheetState.working.knownTable === undefined) return false;
    const { startColumnIndex, endColumnIndex } = this.activeTable;
    return colIndex >= startColumnIndex && colIndex < endColumnIndex;
  }
  gatherFetchRange(gr: SheetGridRangeProps): this {
    this.spreadsheetStateRaw.fetchQueue.gridRanges.push({
      sheetId: this.sheetGid,
      ...gr,
    });
    return this;
  }
  gatherFetchRanges(props: SheetGridRangeProps[]): this {
    props.forEach((props) => this.gatherFetchRange(props));
    return this;
  }
  addSheetChangeToSave(props: SheetChangeProps): this {
    const changes = this.changesToSave;
    switch (props.action) {
      case "sort":
        changes.sort = {
          colIdxToSortBy: props.colIdxToSortBy,
          sortOrder: props.sortOrder,
        };
        break;
      case "insertColumn":
        this._validateColumnInsertAllowed(props.startColumnIndex);
        changes.insertColumn.push(props);
        break;
      case "fill":
        changes.fills.push(Obj.strictOmit(props, "action"));
        break;
      default:
        throw new Error(
          `Invalid action: ${(props as SheetChangeProps).action}. Must be one of "sort", "insertColumn" or "fill".`,
        );
    }
    return this;
  }
  // Where the next Table-end column insert lands: past the inserts already queued.
  get nextEndColumnInsertIndex(): number {
    return (
      this.activeTable.endColumnIndex + this.changesToSave.insertColumn.length
    );
  }
  // Each insert shifts the indexes the next was computed against, so only Table-end inserts may share a flush.
  private _validateColumnInsertAllowed(startColumnIndex: number): void {
    const [firstQueued] = this.changesToSave.insertColumn;
    if (firstQueued === undefined) return;
    if (firstQueued.startColumnIndex !== this.activeTable.endColumnIndex) {
      throw new Error(
        `Refusing to queue a column insert on ${this.sheetLabel}: a mid-Table column insert is already queued.`,
      );
    }
    if (startColumnIndex !== this.nextEndColumnInsertIndex) {
      throw new Error(
        `Refusing to queue a column insert at ${startColumnIndex} on ${this.sheetLabel}: it already has a column insert queued, so the next must land at the Table end, ${this.nextEndColumnInsertIndex}.`,
      );
    }
  }
}
