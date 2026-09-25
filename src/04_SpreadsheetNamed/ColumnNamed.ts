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
import type {
  ColumnIsFormula,
  ColumnName,
  ColumnValue,
  ColumnValueDeclared,
  ColumnValueName,
} from "../01_SpreadsheetSchema/columnConfigsTypes";
import type { SheetName } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import type { VnToCvn } from "../01_SpreadsheetSchema/valueSchemas";
import type { FindReplaceTerms } from "../02_SpreadsheetRaw/ClassTypes/StateRaw";
import type { ColumnRaw } from "../02_SpreadsheetRaw/ColumnRaw";
import type { CellChange } from "../03_SpreadsheetIdentified/ClassTypes/StateIdentified";
import { ColumnIdentified } from "../03_SpreadsheetIdentified/ColumnIdentified";
import { CellNamed } from "./CellNamed";
import { ColumnCommonNamed } from "./ClassBases/ColumnCommonNamed";
import { ColumnMetaNamed } from "./ColumnMetaNamed";
import { SheetNamed } from "./SheetNamed";

export class ColumnNamed<
  SN extends SheetName,
  CN extends ColumnName<SN> = ColumnName<SN>,
> extends ColumnCommonNamed<SN, CN> {
  get sheet(): SheetNamed<SN> {
    return new SheetNamed(this.sheetNamedProps);
  }
  get meta(): ColumnMetaNamed<SN, CN> {
    return new ColumnMetaNamed(this.columnNamedProps);
  }
  get identified(): ColumnIdentified<ColumnValueName<SN, CN>> {
    return new ColumnIdentified<ColumnValueName<SN, CN>>({
      ...this.sheet.identified.sheetIdentifiedProps,
      columnId: this.columnId,
    });
  }
  get raw(): ColumnRaw<VnToCvn<ColumnValueName<SN, CN>>> {
    return this.identified.raw;
  }
  get rowIndexesActive(): number[] {
    return this.identified.cellIndexesActive;
  }
  get valueArrOrEmpty(): ColumnValue<SN, CN>[] {
    return this.identified.valueArrOrEmpty;
  }
  get valueArrFilterEmpty(): NotEmpty<ColumnValue<SN, CN>>[] {
    return this.identified.valueArrFilterEmpty;
  }
  // Not delegated to Identified, so a blank throws with the Named message.
  get valueArrNotEmpty(): NotEmpty<ColumnValue<SN, CN>>[] {
    return this.rowIndexesActive.map((rowIndex) =>
      this.valueNotEmpty(rowIndex),
    );
  }
  get valueArr(): ColumnValueDeclared<SN, CN>[] {
    return this.rowIndexesActive.map((rowIndex) => this.value(rowIndex));
  }
  hasValue(value: ColumnValue<SN, CN>): boolean {
    return this.valueArrOrEmpty.includes(value);
  }
  valueOrEmpty(rowIndex: number): ColumnValue<SN, CN> {
    return this.cell(rowIndex).valueOrEmpty();
  }
  valueNotEmpty(rowIndex: number): NotEmpty<ColumnValue<SN, CN>> {
    return this.cell(rowIndex).valueNotEmpty();
  }
  value(rowIndex: number): ColumnValueDeclared<SN, CN> {
    return this.cell(rowIndex).value();
  }
  cell(rowIndex: number): CellNamed<SN, CN> {
    return new CellNamed({
      ...this.columnNamedProps,
      rowIndex,
    });
  }
  updateAllCells(change: CellChange<ColumnValueName<SN, CN>>): this {
    this.identified.updateAllCells(change);
    return this;
  }
  updateActiveCells(change: CellChange<ColumnValueName<SN, CN>>): this {
    this.identified.updateActiveCells(change);
    return this;
  }
  updateAllFormulas(
    formula: ColumnIsFormula<SN, CN> extends true ? string : never,
  ): this {
    this.identified.updateAllFormulas(formula);
    return this;
  }
  updateActiveFormulas(
    formula: ColumnIsFormula<SN, CN> extends true ? string : never,
  ): this {
    this.identified.updateActiveFormulas(formula);
    return this;
  }
  // Plain strings, unlike every other write here: Google matches the cell's text.
  findReplace(terms: FindReplaceTerms): this {
    this.identified.findReplace(terms);
    return this;
  }
  addConditionalFormatRule(declaration: ConditionalFormatDeclaration): this {
    this.identified.addConditionalFormatRule(declaration);
    return this;
  }
  removeConditionalFormatRules(): this {
    this.identified.removeConditionalFormatRules();
    return this;
  }
  removeConditionalFormatRule(rule: ConditionalFormatRule): this {
    this.identified.removeConditionalFormatRule(rule);
    return this;
  }
  gridRangeFromRow(startRowIndex: number): GridRangeProps {
    return this.identified.gridRangeFromRow(startRowIndex);
  }
  addEditWarning(declaration: EditWarningDeclaration = {}): this {
    this.identified.addEditWarning(declaration);
    return this;
  }
  addEditWarningFromRow(
    startRowIndex: number,
    declaration: EditWarningDeclaration = {},
  ): this {
    this.identified.addEditWarningFromRow(startRowIndex, declaration);
    return this;
  }
  addEditWarningWholeColumn(declaration: EditWarningDeclaration = {}): this {
    this.identified.addEditWarningWholeColumn(declaration);
    return this;
  }
  addEditLock(declaration: EditLockDeclaration = {}): this {
    this.identified.addEditLock(declaration);
    return this;
  }
  addEditLockWholeColumn(declaration: EditLockDeclaration = {}): this {
    this.identified.addEditLockWholeColumn(declaration);
    return this;
  }
  removeEditProtections(): this {
    this.identified.removeEditProtections();
    return this;
  }
  removeEditProtectionsWholeColumn(): this {
    this.identified.removeEditProtectionsWholeColumn();
    return this;
  }
  removeEditProtection(protection: EditProtection): this {
    this.identified.removeEditProtection(protection);
    return this;
  }
  anchoredA1(columnName: ColumnName<SN> = this.columnName): string {
    return this.sheet.column(columnName).identified.anchoredA1();
  }
  prepFetchSpecific(rowIndexes: number[]): this {
    this.identified.prepFetchSpecific(rowIndexes);
    return this;
  }
  prepFetchActive(): this {
    this.identified.prepFetchActive();
    return this;
  }
  prepFetchFull(): this {
    this.identified.prepFetchFull();
    return this;
  }
  activeCellsToDefault(): this {
    this.identified.activeCellsToDefault();
    return this;
  }
  emptyActiveCellsToDefualt(): this {
    this.identified.emptyActiveCellsToDefualt();
    return this;
  }
}
