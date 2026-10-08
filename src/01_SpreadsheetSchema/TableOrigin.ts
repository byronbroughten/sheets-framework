import {
  type SheetColIndex,
  SheetIndex,
  type SheetRowIndex,
} from "../00_Source/RawSource/SheetIndex";
import { type HeadRole, headRows } from "./headRows";

export interface TableOriginProps {
  headerRowIndex: SheetRowIndex;
  startColIndex: SheetColIndex;
}

// Where a Table's header row and first column sit, so its own indexes convert to the sheet's.
export class TableOrigin {
  readonly headerRowIndex: SheetRowIndex;
  readonly startColIndex: SheetColIndex;
  constructor({ headerRowIndex, startColIndex }: TableOriginProps) {
    this.headerRowIndex = headerRowIndex;
    this.startColIndex = startColIndex;
  }
  // Where the framework creates a Table; a managed one is found where its configs record it.
  static expected(): TableOrigin {
    return new TableOrigin({
      headerRowIndex: SheetIndex.row(headRows.countAboveHeader),
      startColIndex: SheetIndex.col(0),
    });
  }
  sheetRowIndex(rowIndex: number): SheetRowIndex {
    return SheetIndex.row(this.headerRowIndex + 1 + rowIndex);
  }
  headSheetRowIndex(role: HeadRole): SheetRowIndex {
    return this.sheetRowIndex(headRows.index(role));
  }
  get topHeadSheetRowIndex(): SheetRowIndex {
    return this.sheetRowIndex(headRows.topIndex);
  }
  sheetColIndex(colIndex: number): SheetColIndex {
    return SheetIndex.col(this.startColIndex + colIndex);
  }
  rowIndex(sheetRowIndex: SheetRowIndex): number {
    return sheetRowIndex - this.headerRowIndex - 1;
  }
  colIndex(sheetColIndex: SheetColIndex): number {
    return sheetColIndex - this.startColIndex;
  }
  // The 1-based number in the sheet's gutter, the one an operator can find.
  rowNumber(rowIndex: number): number {
    return this.headerRowIndex + 2 + rowIndex;
  }
  equals(other: TableOrigin): boolean {
    return (
      this.headerRowIndex === other.headerRowIndex &&
      this.startColIndex === other.startColIndex
    );
  }
}
