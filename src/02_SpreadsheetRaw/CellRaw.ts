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
import type { RowCellChange } from "./ClassTypes/StateRaw";
import { SheetRaw } from "./SheetRaw";

export class CellRaw<
  VN extends CellValueName = CellValueName,
> extends CellBaseRaw {
  get sheet(): SheetRaw {
    return new SheetRaw(this.sheetRawProps);
  }
  get row(): RowCommonRaw {
    return this.sheet.rowCommon(this.rowIndex);
  }
  get gridRange(): BoundedGridRange {
    return {
      sheetId: this.sheetGid,
      startRowIndex: this.rowIndex,
      endRowIndex: this.rowIndex + 1,
      startColumnIndex: this.colIndex,
      endColumnIndex: this.colIndex + 1,
    };
  }
  gatherFetchRange(): this {
    this.sheet.gatherFetchRange(this.gridRange);
    // Sheets omits a never-written cell; finalize treats that as empty.
    const colIndexes =
      this.sheetState.fetchQueue.toFinalize.cells.get(this.rowIndex) ??
      new Set();
    colIndexes.add(this.colIndex);
    this.sheetState.fetchQueue.toFinalize.cells.set(this.rowIndex, colIndexes);
    return this;
  }
  gatherUpdateRequest(change: RowCellChange): void {
    const { formula, ...cellDataChange } = change;
    assertValueAndFormulaExclusive(cellDataChange.value, formula);
    this.updateRequests.update.push({
      kind: "updateCell",
      sheetId: this.sheetGid,
      rowIndex: this.rowIndex,
      colIndex: this.colIndex,
      ...cellDataChange,
      ...(formula !== undefined ? { formula } : {}),
    });
  }
  setValueState(value: CellValue): void {
    if (!this.row.rowIsActive()) {
      throw new Error(
        `Cannot set value for row ${this.rowIndex} because it is not active.`,
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
    if (this.sheet.cellStateIsStale) {
      throw new Error(
        `Cell values went stale when a findReplace was sent; re-fetch before reading row ${this.rowIndex}, column index ${this.colIndex}.`,
      );
    }
    throw new Error(
      `Row ${this.rowIndex} does not have a value set for column index ${this.colIndex}.`,
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
    this.sheet.activeTable.assertRowIndexesNotStale();
    this.sheet.activeTable.validateColIndexNotStale(this.colIndex);
    this.row.validateIsWritable();
    // A row that was never fetched has no state to mirror the write into.
    if (this.row.rowIsActive()) {
      this.setValueState(value);
    }
    this.row.addRowChangeToSave({
      action: "update",
      colIndex: this.colIndex,
      value,
    });
    return this;
  }
  // No state mirror: the next read still sees the old effective value.
  updateFormula(formula: string): this {
    this.sheet.activeTable.assertRowIndexesNotStale();
    validateFormulaString(formula);
    this.sheet.activeTable.validateColIndexNotStale(this.colIndex);
    this.row.validateIsWritable();
    this.row.addRowChangeToSave({
      action: "update",
      colIndex: this.colIndex,
      formula,
    });
    return this;
  }
  // No state mirror: the read path never fetches colour, so there's none to mirror.
  updateBackgroundColor(backgroundColor: RgbColor): this {
    this.sheet.activeTable.assertRowIndexesNotStale();
    this.sheet.activeTable.validateColIndexNotStale(this.colIndex);
    this.row.validateIsWritable();
    this.row.addRowChangeToSave({
      action: "update",
      colIndex: this.colIndex,
      backgroundColor,
    });
    return this;
  }
  addCheckboxValidation(): this {
    this.sheet.activeTable.validateColIndexNotStale(this.colIndex);
    this.sheet.addCheckboxValidationAt(this.gridRange);
    return this;
  }
  addConditionalFormatRule(declaration: ConditionalFormatDeclaration): this {
    this.sheet.addConditionalFormatRuleAt(this.gridRange, declaration);
    return this;
  }
  removeConditionalFormatRules(): this {
    this.sheet.removeConditionalFormatRulesAt(this.gridRange);
    return this;
  }
  removeConditionalFormatRule(rule: ConditionalFormatRule): this {
    this.sheet.removeConditionalFormatRule(rule);
    return this;
  }
  addEditWarning(declaration: EditWarningDeclaration = {}): this {
    this.sheet.addEditWarningAt(this.gridRange, declaration);
    return this;
  }
  addEditLock(declaration: EditLockDeclaration = {}): this {
    this.sheet.addEditLockAt(this.gridRange, declaration);
    return this;
  }
  removeEditProtections(): this {
    this.sheet.removeEditProtectionsAt(this.gridRange);
    return this;
  }
  removeEditProtection(protection: EditProtection): this {
    this.sheet.removeEditProtection(protection);
    return this;
  }
  integrateSnapshot(cell: GridCellSnapshot | undefined): void {
    if (!this.row.rowIsActive()) return;
    this.setValueState(this._queuedValue() ?? cell?.value ?? "");
  }
  // Fills go before per-cell updates in a flush, so a cell's own value wins.
  private _queuedValue(): CellValue | "" | undefined {
    const rowChange = this.sheetState.writeQueue.rows.get(this.rowIndex);
    if (rowChange !== undefined) {
      const cellChange = rowChange.update.get(this.colIndex);
      if (cellChange?.value !== undefined) return cellChange.value;
    }
    const fills = this.sheetState.writeQueue.sheet.fills;
    for (let i = fills.length - 1; i >= 0; i--) {
      const fill = fills[i];
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
    throw new Error("A queued change cannot hold both a value and a formula.");
  }
}
