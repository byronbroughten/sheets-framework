import { Arr } from "../../utils/Arr";
import { Obj } from "../../utils/Obj";
import { ActiveTableRaw } from "../ActiveTableRaw";
import type { SheetGridRangeProps } from "../ClassTypes/AccessorsRaw";
import type {
  ColumnFill,
  RowCellChange,
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
      case "insertTableEndColumn":
        changes.tableEndColumnInsertCount++;
        break;
      case "fill": {
        const fill = Obj.strictOmit(props, "action");
        this._eraseCellFieldsUnder(fill);
        changes.fills.push(fill);
        break;
      }
      default:
        throw new Error(
          `Invalid action: ${(props as SheetChangeProps).action}. Must be one of "sort", "insertTableEndColumn" or "fill".`,
        );
    }
    return this;
  }
  // Fills are sent before per-cell updates, so a later fill wins by erasing what it covers.
  private _eraseCellFieldsUnder(fill: ColumnFill): void {
    for (const [rowIndex, rowChange] of this.sheetState.writeQueue.rows) {
      if (rowIndex < fill.startRowIndex || rowIndex >= fill.endRowIndex) {
        continue;
      }
      const cellChange = rowChange.update.get(fill.colIndex);
      if (cellChange === undefined) continue;
      const fieldsLeft = cellFieldsLeftUnder(fill, cellChange);
      if (Object.keys(fieldsLeft).length === 0) {
        rowChange.update.delete(fill.colIndex);
      } else {
        rowChange.update.set(fill.colIndex, fieldsLeft);
      }
    }
  }
}

function cellFieldsLeftUnder(
  fill: ColumnFill,
  cellChange: RowCellChange,
): RowCellChange {
  const fieldsLeft = { ...cellChange };
  if (fill.value !== undefined || fill.formula !== undefined) {
    delete fieldsLeft.value;
    delete fieldsLeft.formula;
  }
  if (fill.backgroundColor !== undefined) delete fieldsLeft.backgroundColor;
  return fieldsLeft;
}
