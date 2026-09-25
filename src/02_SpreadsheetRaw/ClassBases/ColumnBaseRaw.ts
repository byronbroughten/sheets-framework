import type { ColumnStateRaw } from "../ClassTypes/StateRaw";
import { SheetBaseRaw, type SheetRawProps } from "./SheetBaseRaw";

export interface ColumnRawProps extends SheetRawProps {
  colIndex: number;
}

export class ColumnBaseRaw extends SheetBaseRaw {
  readonly colIndex: number;
  constructor({ colIndex, ...rest }: ColumnRawProps) {
    super(rest);
    this.colIndex = colIndex;
  }
  // Absent until a fetch records a fact about this column.
  get columnState(): ColumnStateRaw | undefined {
    return this.sheetState.working.columnStates.get(this.colIndex);
  }
  get columnRawProps(): ColumnRawProps {
    return {
      colIndex: this.colIndex,
      ...this.sheetRawProps,
    };
  }
}
