// First, out of order: entering the RowCommonRaw <-> TableRaw cycle here loads UniformRowBaseRaw before its base class, in the bundle.
import "./TableRaw";

import type {
  CellValue,
  CellValueName,
} from "../00_Source/CellValues/cellValues";
import type { RowRawProps } from "./ClassBases/RowBaseRaw";
import { RowCommonRaw } from "./ClassBases/RowCommonRaw";

export class RowRaw extends RowCommonRaw {
  constructor(props: RowRawProps) {
    super(props);
    this.validateIsDataRow();
  }
  valueOrEmpty<VN extends CellValueName = CellValueName>(
    colIndex: number,
  ): CellValue<VN> | "" {
    return this.cell<VN>(colIndex).valueOrEmpty();
  }
  get activeValueArr(): CellValue[] {
    return [...this.rowState.values()].map((cellState) => cellState.value);
  }
  private validateIsDataRow(): void {
    if (!this.isDataRow) {
      throw new Error(
        `Cannot perform this operation: ${this.rowLabel(this.rowIndex)} is not a data row.`,
      );
    }
  }
  delete(): void {
    this.table.assertRowIndexesNotStale();
    this.validateTableKeepsADataRow();
    this.remove();
    this.queueRowWrite({ action: "deleteRow" });
    // this.endRowIndex--;
  }
  append(): this {
    if (this.rowIsActive()) {
      throw new Error(
        `Cannot append ${this.rowLabel(this.rowIndex)} because it is already active.`,
      );
    }
    this.tableState.working.rowStates.set(this.rowIndex, new Map());
    this.queueRowWrite({ action: "appendRow" });
    this.table.growDataRowCount();
    return this;
  }
  // A new row copies its formulas from the rows already there, so one must survive.
  private validateTableKeepsADataRow(): void {
    if (!this.table.isDownToLastDataRow) return;
    throw new Error(
      `Cannot delete ${this.rowLabel(this.rowIndex)} of ${this.table.tableLabel}: it is the Table's last data row, and a Table may never be left with none. Clear the row instead.`,
    );
  }
}
