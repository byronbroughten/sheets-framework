import type {
  CellValue,
  CellValueName,
} from "../../00_Source/CellValues/cellValues";
import { Obj } from "../../utils/Obj";
import { CellRaw } from "../CellRaw";
import { emptyStateRaw } from "../ClassTypes/emptyStateRaw";
import type {
  RowWriteFillCellProps,
  RowWriteProps,
  RowWrites,
} from "../ClassTypes/StateRaw";
import { SheetRaw } from "../SheetRaw";
import { RowBaseRaw } from "./RowBaseRaw";

export abstract class RowCommonRaw extends RowBaseRaw {
  get sheet(): SheetRaw {
    return new SheetRaw(this.sheetRawProps);
  }
  // A data row past the table's last row doesn't exist yet — append it instead.
  validateIsWritable(): void {
    super.validateIsWritable();
    if (!this.isDataRow || this.rowIsActive()) return;
    if (this.rowIndex >= this.sheet.activeTable.endRowIndex) {
      throw new Error(
        `Cannot write to row ${this.rowIndex} because it is past the last row of sheetGid ${this.sheetGid}'s table. Append the row first.`,
      );
    }
  }
  // A queued delete outlives a same-run re-fetch; recreating the row would undo it.
  ensureStateExists(): void {
    if (this.isQueuedForDelete) return;
    super.ensureStateExists();
  }
  ensureFullActiveDataCells(): void {
    this.ensureStateExists();
    this.sheet.fullTableColIndexes.forEach((colIndex) => {
      this.cell(colIndex).ensureActive();
    });
  }
  cell<VN extends CellValueName = CellValueName>(
    colIndex: number,
  ): CellRaw<VN> {
    return new CellRaw<VN>({
      ...this.sheetRawProps,
      rowIndex: this.rowIndex,
      colIndex: colIndex,
    });
  }
  abstract get activeValueArr(): CellValue[];
  hasValue(value: unknown): boolean {
    return this.activeValueArr.includes(value as CellValue);
  }
  returnMissingValues<CV extends CellValue>(...values: CV[]): CV[] {
    return values.filter((value) => !this.activeValueArr.includes(value));
  }
  remove(): void {
    this.rowStates.delete(this.rowIndex);
  }
  updateValue(colIndex: number, value: CellValue): this {
    this.cell(colIndex).updateValue(value);
    return this;
  }
  gatherFetchFull(): this {
    this.sheet.gatherFetchRange({
      startRowIndex: this.rowIndex,
      endRowIndex: this.rowIndex + 1,
      startColumnIndex: this.sheet.activeTable.startColumnIndex,
    });
    this.sheetState.fetchQueue.toFinalize.rows.add(this.rowIndex);
    return this;
  }
  get isQueuedForDelete(): boolean {
    return (
      this.sheetState.writeQueue.rows.get(this.rowIndex)?.deleteRow === true
    );
  }
  get writes(): RowWrites {
    this._ensureWritesExist();
    return this.sheetState.writeQueue.rows.get(this.rowIndex) as RowWrites;
  }
  private _ensureWritesExist(): void {
    const rowWrites = this.sheetState.writeQueue.rows;
    if (!rowWrites.has(this.rowIndex)) {
      rowWrites.set(this.rowIndex, emptyStateRaw.rowWrites());
    }
  }
  queueRowWrite(props: RowWriteProps): this {
    const writes = this.writes;
    if (writes.deleteRow) return this;
    const actions = {
      appendRow: (_: RowWriteProps) => (writes.appendRow = true),
      deleteRow: (_: RowWriteProps) => (writes.deleteRow = true),
      fillCell: (props: RowWriteProps) => {
        const { colIndex, ...rest } = props as RowWriteFillCellProps;
        const incoming = Obj.strictOmit(rest, "action");
        const merged = {
          ...writes.fillCells.get(colIndex),
          ...incoming,
        };
        if ("formula" in incoming) delete merged.value;
        if ("value" in incoming) delete merged.formula;
        writes.fillCells.set(colIndex, merged);
      },
    };
    actions[props.action](props);
    return this;
  }
  gatherAppendRowsOperation(): void {
    // One request per table: Sheets treats each appendCells as targeting the
    // same first free row, so N one-row requests only grow the table by one.
    const existing = this.writeOperations.appendRows.find(
      (operation) => operation.sheetId === this.sheetGid,
    );
    if (existing) {
      existing.emptyRowCount += 1;
      return;
    }
    this.writeOperations.appendRows.push({
      kind: "appendRows",
      sheetId: this.sheetGid,
      tableId: `${this.sheet.activeTable.tableId}`,
      emptyRowCount: 1,
    });
  }
}
