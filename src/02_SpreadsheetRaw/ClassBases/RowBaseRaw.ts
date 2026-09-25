import type { CellValue } from "../../00_Source/CellValues/cellValues";
import type { RowStateRaw } from "../ClassTypes/StateRaw";
import { SheetBaseRaw, type SheetRawProps } from "./SheetBaseRaw";

export interface RowRawProps extends SheetRawProps {
  rowIndex: number;
}

export class RowBaseRaw extends SheetBaseRaw {
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
    return this.rowIndex >= this.schema.topDataRowIdx;
  }
  get rowState(): RowStateRaw {
    return this.getRowState(this.rowIndex);
  }
  rowIsActive(): boolean {
    return this.sheetState.working.rowStates.has(this.rowIndex);
  }
  get isReserved(): boolean {
    return this.sheetState.writeQueue.reservedRowIndexes.has(this.rowIndex);
  }
  reserve(): void {
    this.sheetState.writeQueue.reservedRowIndexes.add(this.rowIndex);
  }
  release(): void {
    this.sheetState.writeQueue.reservedRowIndexes.delete(this.rowIndex);
  }
  validateIsWritable(): void {
    if (!this.isDataRow || this.rowIsActive()) return;
    if (this.sheetState.working.knownTable === undefined) {
      throw new Error(
        `Cannot write to row ${this.rowIndex} of sheetGid ${this.sheetGid} before its sheet properties have been fetched.`,
      );
    }
  }
  validateIsActive(): void {
    if (!this.rowIsActive()) {
      throw new Error(
        `Row ${this.rowIndex} is not active. Cannot perform this operation.`,
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
      `Value ${value} not found in row ${this.rowIndex}. Cannot find column index.`,
    );
  }
  get rowRawProps(): RowRawProps {
    return {
      ...this.sheetRawProps,
      rowIndex: this.rowIndex,
    };
  }
}
