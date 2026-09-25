import type { NotEmpty } from "../00_Source/CellValues/cellValues";
import type {
  ConditionalFormatDeclaration,
  ConditionalFormatRule,
} from "../00_Source/RawSource/ConditionalFormat";
import type {
  EditLockDeclaration,
  EditProtection,
  EditWarningDeclaration,
} from "../00_Source/RawSource/EditProtection";
import type { GridRangeProps } from "../00_Source/RawSource/RawSource";
import {
  toWireValue,
  type Value,
  type ValueName,
  type VnToCvn,
} from "../01_SpreadsheetSchema/valueSchemas";
import type {
  FindReplaceTerms,
  RowCellChange,
} from "../02_SpreadsheetRaw/ClassTypes/StateRaw";
import { ColumnRaw } from "../02_SpreadsheetRaw/ColumnRaw";
import { CellIdentified } from "./CellIdentified";
import { ColumnCommonIdentified } from "./ClassBases/ColumnCommonIdentified";
import type { CellChange } from "./ClassTypes/StateIdentified";
import { ColumnMetaIdentified } from "./ColumnMetaIdentified";
import { SheetIdentified } from "./SheetIdentified";

export class ColumnIdentified<
  VN extends ValueName = ValueName,
> extends ColumnCommonIdentified<VN> {
  get sheet(): SheetIdentified {
    return new SheetIdentified(this.sheetIdentifiedProps);
  }
  get meta(): ColumnMetaIdentified<VN> {
    return new ColumnMetaIdentified(this.columnIdentifiedProps);
  }
  get raw(): ColumnRaw<VnToCvn<VN>> {
    return new ColumnRaw({
      ...this.sheetIdentifiedProps,
      colIndex: this.colIndex,
    });
  }
  get cellIndexesActive(): number[] {
    return this.raw.cellIndexesActive;
  }
  get cellIndexesFull(): number[] {
    return this.raw.cellIndexesFull;
  }
  get cellsFull(): CellIdentified<VN>[] {
    return this.cellIndexesFull.map((rowIndex) => this.cell(rowIndex));
  }
  prepFetchSpecific(rowIndexes: number[]): this {
    rowIndexes.forEach((rowIndex) => {
      this.cell(rowIndex).prepFetch();
    });
    return this;
  }
  prepFetchActive(): this {
    return this.prepFetchSpecific(this.cellIndexesActive);
  }
  prepFetchFull(): this {
    this.fetchTargets.push({
      kind: "fullDataColumn",
      column: this.columnId,
    });
    return this;
  }
  // Through the cells, not straight to Raw, so the value name's blank is read here too.
  get valueArrOrEmpty(): Value<VN>[] {
    return this.sheet.rowIndexesActive.map((rowIndex) =>
      this.valueOrEmpty(rowIndex),
    );
  }
  get valueArrFilterEmpty(): NotEmpty<Value<VN>>[] {
    return this.valueArrOrEmpty.filter(
      (value): value is NotEmpty<Value<VN>> => value !== "",
    );
  }
  get valueArrNotEmpty(): NotEmpty<Value<VN>>[] {
    return this.sheet.rowIndexesActive.map((rowIndex) =>
      this.cell(rowIndex).valueNotEmpty(),
    );
  }
  hasValue(value: Value<VN>): boolean {
    return this.valueArrOrEmpty.includes(value);
  }
  valueOrEmpty(rowIndex: number): Value<VN> {
    return this.cell(rowIndex).valueOrEmpty();
  }
  valueNotEmpty(rowIndex: number): NotEmpty<Value<VN>> {
    return this.cell(rowIndex).valueNotEmpty();
  }
  cell(rowIndex: number): CellIdentified<VN> {
    return new CellIdentified({
      ...this.columnIdentifiedProps,
      rowIndex,
    });
  }
  get cellsActive(): CellIdentified<VN>[] {
    return this.cellIndexesActive.map((rowIndex) => this.cell(rowIndex));
  }
  activeCellsToDefault(): void {
    this.cellsActive.forEach((cell) => {
      cell.updateToDefault();
    });
  }
  allCellsToDefault(): void {
    this.cellsFull.forEach((cell) => {
      cell.updateToDefault();
    });
  }
  updateAllCells(change: CellChange<VN>): this {
    this.raw.updateAllCells(this._rawChange(change));
    return this;
  }
  updateActiveCells(change: CellChange<VN>): this {
    this.raw.updateActiveCells(this._rawChange(change));
    return this;
  }
  updateAllFormulas(formula: string): this {
    this.schema.validateIsFormula();
    this.raw.updateAllFormulas(formula);
    return this;
  }
  updateActiveFormulas(formula: string): this {
    this.schema.validateIsFormula();
    this.raw.updateActiveFormulas(formula);
    return this;
  }
  // Google matches the text, so neither string is checked against the value config.
  findReplace(terms: FindReplaceTerms): this {
    this.raw.findReplace(terms);
    return this;
  }
  addConditionalFormatRule(declaration: ConditionalFormatDeclaration): this {
    this.raw.addConditionalFormatRule(declaration);
    return this;
  }
  removeConditionalFormatRules(): this {
    this.raw.removeConditionalFormatRules();
    return this;
  }
  removeConditionalFormatRule(rule: ConditionalFormatRule): this {
    this.raw.removeConditionalFormatRule(rule);
    return this;
  }
  gridRangeFromRow(startRowIndex: number): GridRangeProps {
    return this.raw.gridRangeFromRow(startRowIndex);
  }
  addEditWarning(declaration: EditWarningDeclaration = {}): this {
    this.raw.addEditWarning(declaration);
    return this;
  }
  addEditWarningFromRow(
    startRowIndex: number,
    declaration: EditWarningDeclaration = {},
  ): this {
    this.raw.addEditWarningFromRow(startRowIndex, declaration);
    return this;
  }
  addEditWarningWholeColumn(declaration: EditWarningDeclaration = {}): this {
    this.raw.addEditWarningWholeColumn(declaration);
    return this;
  }
  addEditLock(declaration: EditLockDeclaration = {}): this {
    this.raw.addEditLock(declaration);
    return this;
  }
  addEditLockWholeColumn(declaration: EditLockDeclaration = {}): this {
    this.raw.addEditLockWholeColumn(declaration);
    return this;
  }
  removeEditProtections(): this {
    this.raw.removeEditProtections();
    return this;
  }
  removeEditProtectionsWholeColumn(): this {
    this.raw.removeEditProtectionsWholeColumn();
    return this;
  }
  removeEditProtection(protection: EditProtection): this {
    this.raw.removeEditProtection(protection);
    return this;
  }
  anchoredA1(colIndex = this.colIndex): string {
    return this.sheet.anchoredA1(colIndex);
  }
  // A colour-only write is legitimate on a formula column; a value is not.
  private _rawChange({
    value,
    ...rest
  }: CellChange<VN>): RowCellChange<VnToCvn<VN>> {
    if (value === undefined) return rest;
    this.schema.validateDataNotFormula();
    return { ...rest, value: toWireValue(value) };
  }
  emptyActiveCellsToDefualt(): this {
    this.cellsActive.forEach((cell) => {
      if (cell.raw.isEmpty) {
        cell.updateToDefault();
      }
    });
    return this;
  }
}
