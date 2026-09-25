import { Val } from "../../utils/Val";
import type { CellStateRaw, RowStateRaw } from "../ClassTypes/StateRaw";
import { ColumnBaseRaw, type ColumnRawProps } from "./ColumnBaseRaw";

export interface CellRawProps extends ColumnRawProps {
  rowIndex: number;
}

export class CellBaseRaw extends ColumnBaseRaw {
  readonly rowIndex: number;
  constructor({ rowIndex, ...rest }: CellRawProps) {
    super(rest);
    this.rowIndex = rowIndex;
  }
  get rowState(): RowStateRaw {
    return this.getRowState(this.rowIndex);
  }
  get cellState(): CellStateRaw {
    return Val.assert(
      this.rowState.get(this.colIndex),
      `cellState for row ${this.rowIndex}, column ${this.colIndex} on sheetGid ${this.sheetGid}`,
    );
  }
  get cellRawProps(): CellRawProps {
    return {
      rowIndex: this.rowIndex,
      ...this.columnRawProps,
    };
  }
}
