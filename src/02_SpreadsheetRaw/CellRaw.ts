import type {
  CellValue,
  CellValueName,
} from "../00_Source/CellValues/cellValues";
import type {
  ConditionalFormatDeclaration,
  ConditionalFormatRule,
} from "../00_Source/RawSource/ConditionalFormat";
import type {
  EditLockDeclaration,
  EditProtection,
  EditWarningDeclaration,
} from "../00_Source/RawSource/EditProtection";
import type {
  BoundedGridRange,
  GridCellSnapshot,
} from "../00_Source/RawSource/RawSource";
import type { RgbColor } from "../00_Source/RawSource/RgbColor";
import { CellBaseRaw } from "./ClassBases/CellBaseRaw";
import type { RowCommonRaw } from "./ClassBases/RowCommonRaw";
import type { CellFill } from "./ClassTypes/StateRaw";
import { TableRaw } from "./TableRaw";

export class CellRaw<
  VN extends CellValueName = CellValueName,
> extends CellBaseRaw {
  get table(): TableRaw {
    return new TableRaw(this.tableRawProps);
  }
  get row(): RowCommonRaw {
    return this.table.rowCommon(this.rowIndex);
  }
  gridRange(): BoundedGridRange {
    const origin = this.tableOrigin();
    return {
      sheetId: this.sheetGid,
      startRowIndex: origin.sheetRowIndex(this.rowIndex),
      endRowIndex: origin.sheetRowIndex(this.rowIndex + 1),
      startColumnIndex: origin.sheetColIndex(this.colIndex),
      endColumnIndex: origin.sheetColIndex(this.colIndex + 1),
    };
  }
  gatherFetchRange(): this {
    this.table.gatherFetchRange(this.gridRange());
    return this.prepFetchBackfill();
  }
  // Marks the cell for finalize's blank backfill without gathering its own range.
  prepFetchBackfill(): this {
    // Sheets omits a never-written cell; finalize treats that as empty.
    const colIndexes =
      this.tableState.fetchQueue.toFinalize.cells.get(this.rowIndex) ??
      new Set();
    colIndexes.add(this.colIndex);
    this.tableState.fetchQueue.toFinalize.cells.set(this.rowIndex, colIndexes);
    return this;
  }
  gatherFillCellOperation(cellFill: CellFill): void {
    const { formula, ...cellData } = cellFill;
    assertValueAndFormulaExclusive(cellData.value, formula);
    const origin = this.tableOrigin();
    this.writeOperations.fillCell.push({
      kind: "fillCell",
      sheetId: this.sheetGid,
      rowIndex: origin.sheetRowIndex(this.rowIndex),
      colIndex: origin.sheetColIndex(this.colIndex),
      ...cellData,
      ...(formula !== undefined ? { formula } : {}),
    });
  }
  setValueState(value: CellValue): void {
    if (!this.row.rowIsActive()) {
      throw new Error(
        `Cannot set value for ${this.rowLabel(this.rowIndex)} because it is not active.`,
      );
    }
    this.rowState.set(this.colIndex, { value });
  }
  get isEmpty(): boolean {
    this.validateIsActive();
    return this.cellState.value === "";
  }
  validateIsActive(): void {
    if (this.isActive) return;
    if (this.table.cellStateIsStale) {
      throw new Error(
        `Cell values went stale when a findReplace was sent; re-fetch before reading ${this.rowLabel(this.rowIndex)}, column index ${this.colIndex}.`,
      );
    }
    throw new Error(
      `No value is set in ${this.rowLabel(this.rowIndex)} for column index ${this.colIndex}.`,
    );
  }
  get isActive(): boolean {
    return this.row.rowIsActive() && this.rowState.has(this.colIndex);
  }
  ensureActive(): void {
    if (!this.row.rowIsActive()) return;
    if (!this.isActive) {
      this.setValueState("");
    }
  }
  // An untouched cell holds nothing; Raw reports that rather than judging it.
  valueOrEmpty(): CellValue<VN> | "" {
    this.validateIsActive();
    return this.cellState.value as CellValue<VN> | "";
  }
  updateValue(value: CellValue<VN>): this {
    this.table.assertRowIndexesNotStale();
    this.row.validateIsWritable();
    // A row that was never fetched has no state to mirror the write into.
    if (this.row.rowIsActive()) {
      this.setValueState(value);
    }
    this.row.queueRowWrite({
      action: "fillCell",
      colIndex: this.colIndex,
      value,
    });
    return this;
  }
  // No state mirror: the next read still sees the old effective value.
  updateFormula(formula: string): this {
    this.table.assertRowIndexesNotStale();
    validateFormulaString(formula);
    this.row.validateIsWritable();
    this.row.queueRowWrite({
      action: "fillCell",
      colIndex: this.colIndex,
      formula,
    });
    return this;
  }
  // No state mirror: the read path never fetches colour, so there's none to mirror.
  updateBackgroundColor(backgroundColor: RgbColor): this {
    this.table.assertRowIndexesNotStale();
    this.row.validateIsWritable();
    this.row.queueRowWrite({
      action: "fillCell",
      colIndex: this.colIndex,
      backgroundColor,
    });
    return this;
  }
  addCheckboxValidation(): this {
    this.table.addCheckboxValidationAt(this.gridRange());
    return this;
  }
  addConditionalFormatRule(declaration: ConditionalFormatDeclaration): this {
    this.table.addConditionalFormatRuleAt(this.gridRange(), declaration);
    return this;
  }
  removeConditionalFormatRules(): this {
    this.table.removeConditionalFormatRulesAt(this.gridRange());
    return this;
  }
  removeConditionalFormatRule(rule: ConditionalFormatRule): this {
    this.table.removeConditionalFormatRule(rule);
    return this;
  }
  addEditWarning(declaration: EditWarningDeclaration = {}): this {
    this.table.addEditWarningAt(this.gridRange(), declaration);
    return this;
  }
  addEditLock(declaration: EditLockDeclaration = {}): this {
    this.table.addEditLockAt(this.gridRange(), declaration);
    return this;
  }
  removeEditProtections(): this {
    this.table.removeEditProtectionsAt(this.gridRange());
    return this;
  }
  removeEditProtection(protection: EditProtection): this {
    this.table.removeEditProtection(protection);
    return this;
  }
  integrateSnapshot(cell: GridCellSnapshot | undefined): void {
    if (!this.row.rowIsActive()) return;
    this.setValueState(this._queuedValue() ?? cell?.value ?? "");
  }
  // A fill erases the cell writes queued before it, so a cell's own value is the latest.
  private _queuedValue(): CellValue | "" | undefined {
    const rowWrites = this.tableState.writeQueue.rows.get(this.rowIndex);
    if (rowWrites !== undefined) {
      const cellFill = rowWrites.fillCells.get(this.colIndex);
      if (cellFill?.value !== undefined) return cellFill.value;
    }
    const fillColumns = this.tableState.writeQueue.table.fillColumns;
    for (let i = fillColumns.length - 1; i >= 0; i--) {
      const fill = fillColumns[i];
      if (fill === undefined || fill.value === undefined) continue;
      if (fill.colIndex !== this.colIndex) continue;
      if (
        this.rowIndex < fill.startRowIndex ||
        this.rowIndex >= fill.endRowIndex
      ) {
        continue;
      }
      return fill.value;
    }
    return undefined;
  }
}

export function validateFormulaString(formula: string): void {
  if (formula.startsWith("=")) return;
  throw new Error(`Formula must start with "=". Got "${formula}".`);
}

export function assertValueAndFormulaExclusive(
  value: unknown,
  formula: string | undefined,
): void {
  if (formula !== undefined && value !== undefined) {
    throw new Error("A queued write cannot hold both a value and a formula.");
  }
}
