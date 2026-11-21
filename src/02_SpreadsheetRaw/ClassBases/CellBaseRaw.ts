import { Val } from "@byronbroughten/utils/val";

import type { CellStateRaw, RowStateRaw } from "../ClassTypes/StateRaw.ts";
import { ColumnBaseRaw, type ColumnRawProps } from "./ColumnBaseRaw.ts";

export type CellRawProps = ColumnRawProps & { rowIndex: number };

export class CellBaseRaw extends ColumnBaseRaw {
  declare readonly rowIndex: number;
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
      `cellState for ${this.rowLabel(this.rowIndex)}, column index ${this.colIndex} on sheetGid ${this.sheetGid}`,
    );
  }
}
