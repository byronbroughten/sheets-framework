import type { ColumnStateRaw } from "../ClassTypes/StateRaw";
import { TableBaseRaw, type TableRawProps } from "./TableBaseRaw";

export type ColumnRawProps = TableRawProps & { colIndex: number };

export class ColumnBaseRaw extends TableBaseRaw {
  readonly colIndex: number;
  constructor({ colIndex, ...rest }: ColumnRawProps) {
    super(rest);
    this.colIndex = colIndex;
  }
  // Absent until a fetch records a fact about this column.
  get columnState(): ColumnStateRaw | undefined {
    return this.tableState.working.columnStates.get(this.colIndex);
  }
  get columnRawProps(): ColumnRawProps {
    return {
      colIndex: this.colIndex,
      ...this.tableRawProps,
    };
  }
}
