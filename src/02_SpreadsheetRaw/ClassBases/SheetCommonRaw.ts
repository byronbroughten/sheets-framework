import { Arr } from "../../utils/Arr";
import { Obj } from "../../utils/Obj";
import { ActiveTableRaw } from "../ActiveTableRaw";
import type { SheetGridRangeProps } from "../ClassTypes/AccessorsRaw";
import type {
  CellFill,
  ColumnFill,
  SheetWriteProps,
  SheetWrites,
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
    if (knownTable.endRowIndex <= knownTable.startRowIndex + 1) {
      throw new Error(
        `Sheet ${this.sheetLabel} Table must have at least one data row.`,
      );
    }
    return new ActiveTableRaw(this.sheetRawProps);
  }
  get fullTableColIndexes(): number[] {
    return Arr.indexesFromUntil(0, this.activeTable.columnCount);
  }
  get writes(): SheetWrites {
    return this.sheetState.writeQueue.sheet;
  }
  // The table's own range, not the layout's: no table means no table columns.
  isTableColIndex(colIndex: number): boolean {
    if (this.sheetState.working.knownTable === undefined) return false;
    return colIndex >= 0 && colIndex < this.activeTable.columnCount;
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
  queueSheetWrite(props: SheetWriteProps): this {
    const writes = this.writes;
    switch (props.action) {
      case "sort":
        writes.sort = {
          colIdxToSortBy: props.colIdxToSortBy,
          sortOrder: props.sortOrder,
        };
        break;
      case "insertTableEndColumn":
        writes.insertTableEndColumnCount++;
        break;
      case "fillColumn": {
        const fill = Obj.strictOmit(props, "action");
        this._eraseCellFieldsUnder(fill);
        writes.fillColumns.push(fill);
        break;
      }
      default:
        throw new Error(
          `Invalid action: ${(props as SheetWriteProps).action}. Must be one of "sort", "insertTableEndColumn" or "fillColumn".`,
        );
    }
    return this;
  }
  // Column fills are sent before cell writes, so a later column fill wins by erasing what it covers.
  private _eraseCellFieldsUnder(fill: ColumnFill): void {
    for (const [rowIndex, rowWrites] of this.sheetState.writeQueue.rows) {
      if (rowIndex < fill.startRowIndex || rowIndex >= fill.endRowIndex) {
        continue;
      }
      const cellFill = rowWrites.fillCells.get(fill.colIndex);
      if (cellFill === undefined) continue;
      const fieldsLeft = cellFieldsLeftUnder(fill, cellFill);
      if (Object.keys(fieldsLeft).length === 0) {
        rowWrites.fillCells.delete(fill.colIndex);
      } else {
        rowWrites.fillCells.set(fill.colIndex, fieldsLeft);
      }
    }
  }
}

function cellFieldsLeftUnder(fill: ColumnFill, cellFill: CellFill): CellFill {
  const fieldsLeft = { ...cellFill };
  if (fill.value !== undefined || fill.formula !== undefined) {
    delete fieldsLeft.value;
    delete fieldsLeft.formula;
  }
  if (fill.backgroundColor !== undefined) delete fieldsLeft.backgroundColor;
  return fieldsLeft;
}
