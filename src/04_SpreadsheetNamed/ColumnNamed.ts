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
import type {
  GridRangeProps,
  TableColumnType,
} from "../00_Source/RawSource/RawSource";
import type {
  ColumnIsFormula,
  ColumnName,
  ColumnValue,
  ColumnValueDeclared,
  ColumnValueName,
} from "../01_SpreadsheetSchema/columnConfigsTypes";
import type {
  HeadRole,
  HeadRowValueName,
} from "../01_SpreadsheetSchema/headRows";
import type { TableName } from "../01_SpreadsheetSchema/tableConfigsTypes";
import type { VnToCvn } from "../01_SpreadsheetSchema/valueSchemas";
import type { FindReplaceTerms } from "../02_SpreadsheetRaw/ClassTypes/StateRaw";
import type { ColumnRaw } from "../02_SpreadsheetRaw/ColumnRaw";
import type { CellIdentified } from "../03_SpreadsheetIdentified/CellIdentified";
import type { CellChange } from "../03_SpreadsheetIdentified/ClassTypes/StateIdentified";
import { ColumnIdentified } from "../03_SpreadsheetIdentified/ColumnIdentified";
import { CellNamed } from "./CellNamed";
import { ColumnCommonNamed } from "./ClassBases/ColumnCommonNamed";
import { ColumnMetaNamed } from "./ColumnMetaNamed";
import { TableNamed } from "./TableNamed";

export class ColumnNamed<
  TN extends TableName,
  CN extends ColumnName<TN> = ColumnName<TN>,
> extends ColumnCommonNamed<TN, CN> {
  get table(): TableNamed<TN> {
    return new TableNamed(this.sheetNamedProps);
  }
  get meta(): ColumnMetaNamed<TN, CN> {
    return new ColumnMetaNamed(this.columnNamedProps);
  }
  get identified(): ColumnIdentified<ColumnValueName<TN, CN>> {
    return new ColumnIdentified<ColumnValueName<TN, CN>>({
      ...this.table.identified.tableIdentifiedProps,
      columnId: this.columnId,
    });
  }
  get raw(): ColumnRaw<VnToCvn<ColumnValueName<TN, CN>>> {
    return this.identified.raw;
  }
  get workingRowIndexes(): number[] {
    return this.identified.workingCellIndexes;
  }
  get valueArrOrEmpty(): ColumnValue<TN, CN>[] {
    return this.identified.valueArrOrEmpty;
  }
  get valueArrFilterEmpty(): NotEmpty<ColumnValue<TN, CN>>[] {
    return this.identified.valueArrFilterEmpty;
  }
  // Not delegated to Identified, so a blank throws with the Named message.
  get valueArrNotEmpty(): NotEmpty<ColumnValue<TN, CN>>[] {
    return this.workingRowIndexes.map((rowIndex) =>
      this.valueNotEmpty(rowIndex),
    );
  }
  get valueArr(): ColumnValueDeclared<TN, CN>[] {
    return this.workingRowIndexes.map((rowIndex) => this.value(rowIndex));
  }
  // Identified, not Named: a Named cell's value type is its column's, which a head cell doesn't share.
  headCell<HR extends HeadRole>(
    headRole: HR,
  ): CellIdentified<HeadRowValueName<HR>> {
    return this.identified.headCell(headRole);
  }
  hasValue(value: ColumnValue<TN, CN>): boolean {
    return this.valueArrOrEmpty.includes(value);
  }
  valueOrEmpty(rowIndex: number): ColumnValue<TN, CN> {
    return this.cell(rowIndex).valueOrEmpty();
  }
  valueNotEmpty(rowIndex: number): NotEmpty<ColumnValue<TN, CN>> {
    return this.cell(rowIndex).valueNotEmpty();
  }
  value(rowIndex: number): ColumnValueDeclared<TN, CN> {
    return this.cell(rowIndex).value();
  }
  cell(rowIndex: number): CellNamed<TN, CN> {
    return new CellNamed({
      ...this.columnNamedProps,
      rowIndex,
    });
  }
  updateAllCells(change: CellChange<ColumnValueName<TN, CN>>): this {
    this.identified.updateAllCells(change);
    return this;
  }
  updateWorkingCells(change: CellChange<ColumnValueName<TN, CN>>): this {
    this.identified.updateWorkingCells(change);
    return this;
  }
  updateAllFormulas(
    formula: ColumnIsFormula<TN, CN> extends true ? string : never,
  ): this {
    this.identified.updateAllFormulas(formula);
    return this;
  }
  updateWorkingFormulas(
    formula: ColumnIsFormula<TN, CN> extends true ? string : never,
  ): this {
    this.identified.updateWorkingFormulas(formula);
    return this;
  }
  updateColumnType(columnType: TableColumnType): this {
    this.identified.updateColumnType(columnType);
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
  anchoredA1(columnName: ColumnName<TN> = this.columnName): string {
    return this.table.column(columnName).identified.anchoredA1();
  }
  prepFetchSpecific(rowIndexes: number[]): this {
    this.identified.prepFetchSpecific(rowIndexes);
    return this;
  }
  prepFetchWorking(): this {
    this.identified.prepFetchWorking();
    return this;
  }
  prepFetchFull(): this {
    this.identified.prepFetchFull();
    return this;
  }
  workingCellsToDefault(): this {
    this.identified.workingCellsToDefault();
    return this;
  }
  emptyWorkingCellsToDefault(): this {
    this.identified.emptyWorkingCellsToDefault();
    return this;
  }
}
