import type { CellValue } from "../../00_Source/CellValues/cellValues";
import type { RowStateRaw } from "../ClassTypes/StateRaw";
import { TableBaseRaw, type TableRawProps } from "./TableBaseRaw";

export type RowRawProps = TableRawProps & { rowIndex: number };

export class RowBaseRaw extends TableBaseRaw {
  readonly rowIndex;
  constructor({ rowIndex, ...rest }: RowRawProps) {
    super(rest);
    this.rowIndex = rowIndex;
  }
  ensureStateExists(): void {
    if (!this.rowIsActive()) {
      this.rowStates.set(this.rowIndex, new Map());
    }
  }
  get isDataRow(): boolean {
    return this.rowIndex >= 0;
  }
  get rowState(): RowStateRaw {
    return this.getRowState(this.rowIndex);
  }
  rowIsActive(): boolean {
    return this.tableState.working.rowStates.has(this.rowIndex);
  }
  get isReserved(): boolean {
    return this.tableState.writeQueue.reservedRowIndexes.has(this.rowIndex);
  }
  reserve(): void {
    this.tableState.writeQueue.reservedRowIndexes.add(this.rowIndex);
  }
  release(): void {
    this.tableState.writeQueue.reservedRowIndexes.delete(this.rowIndex);
  }
  validateIsWritable(): void {
    if (!this.isDataRow || this.rowIsActive()) return;
    if (this.tableProperties === undefined) {
      throw new Error(
        `Cannot write to ${this.rowLabel(this.rowIndex)} of sheetGid ${this.sheetGid} before its sheet properties have been fetched.`,
      );
    }
  }
  validateIsActive(): void {
    if (!this.rowIsActive()) {
      throw new Error(
        `Cannot perform this operation: ${this.rowLabel(this.rowIndex)} is not active.`,
      );
    }
  }
  colIndexOfValue(value: CellValue): number {
    for (const [colIndex, cellState] of this.rowState.entries()) {
      if (cellState.value === value) {
        return colIndex;
      }
    }
    throw new Error(
      `Value ${value} not found in ${this.rowLabel(this.rowIndex)}. Cannot find column index.`,
    );
  }
  get rowRawProps(): RowRawProps {
    return {
      ...this.tableRawProps,
      rowIndex: this.rowIndex,
    };
  }
}
