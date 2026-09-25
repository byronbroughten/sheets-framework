import type {
  ConditionalFormatDeclaration,
  ConditionalFormatRule,
} from "../00_Source/RawSource/ConditionalFormat";
import type {
  EditLockDeclaration,
  EditProtection,
  EditWarningDeclaration,
  WholeSheetEditLockDeclaration,
  WholeSheetEditWarningDeclaration,
} from "../00_Source/RawSource/EditProtection";
import type {
  ColumnName,
  ColumnValue,
  SheetDataValues,
  SheetDataValuesAll,
} from "../01_SpreadsheetSchema/columnConfigsTypes";
import type { SheetName } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import type { FindReplaceTerms } from "../02_SpreadsheetRaw/ClassTypes/StateRaw";
import type { SheetRaw } from "../02_SpreadsheetRaw/SheetRaw";
import { ColumnIdentified } from "../03_SpreadsheetIdentified/ColumnIdentified";
import { SheetIdentified } from "../03_SpreadsheetIdentified/SheetIdentified";
import { Arr } from "../utils/Arr";
import { Obj } from "../utils/Obj";
import { Val } from "../utils/Val";
import { SheetCommonNamed } from "./ClassBases/SheetCommonNamed";
import { ColumnNamed } from "./ColumnNamed";
import { RowNamed } from "./RowNamed";
import { SheetMetaNamed } from "./SheetMetaNamed";
import type { SheetNameWithIdAndNameColumn } from "./SheetNameGroups";
import type { RowIdByName } from "./Types/RowIdByName";

/**
 * Name-addressed primary sheet: data rows, append, named columns.
 * Structure as its own shape is this.meta (SheetMetaNamed).
 * Crossing views costs one word: meta / primary; no cross-level shortcut.
 * Member placement for Meta vs primary: docs/vocabulary.md, "Meta / primary".
 * docs/architecture/class-chains.md
 */
export class SheetNamed<
  SN extends SheetName = SheetName,
> extends SheetCommonNamed<SN> {
  get meta(): SheetMetaNamed<SN> {
    return new SheetMetaNamed(this.sheetNamedProps);
  }
  get raw(): SheetRaw {
    return this.identified.raw;
  }
  get identified(): SheetIdentified {
    return new SheetIdentified({
      ...this.sheetNamedProps,
      sheetGid: this.sheetGid,
    });
  }
  get rowIndexesActive(): number[] {
    return this.identified.rowIndexesActive;
  }
  get rowIndexesActiveWithData(): number[] {
    return this.identified.rowIndexesActiveWithData;
  }
  get rowIndexesFullWithData(): number[] {
    return this.identified.rowIndexesFullWithData;
  }
  get rows(): RowNamed<SN>[] {
    return this.identified.rows.map((row) => this.row(row.rowIndex));
  }
  get topRow(): RowNamed<SN> {
    return this.row(this.schema.topDataRowIdx);
  }
  row(rowIndex: number): RowNamed<SN> {
    return new RowNamed({
      ...this.sheetNamedProps,
      rowIndex,
    });
  }
  column<CN extends ColumnName<SN>>(columnName: CN): ColumnNamed<SN, CN> {
    return new ColumnNamed({
      ...this.sheetNamedProps,
      columnName,
    });
  }
  // By id, so the column name's value type isn't composed into the result.
  columnIdentified(columnName: ColumnName<SN>): ColumnIdentified {
    const { columnId } = this.schema.columnByName(columnName);
    return this.identified.column(columnId);
  }
  columns<CS extends readonly ColumnName<SN>[]>(
    ...columnNames: CS
  ): { [K in CS[number]]: ColumnNamed<SN, K> } {
    const columns = {} as { [K in CS[number]]: ColumnNamed<SN, K> };
    columnNames.forEach((columnName) => {
      columns[columnName] = this.column(columnName);
    });
    return columns;
  }
  prepFetchColumnsFull<CS extends readonly ColumnName<SN>[]>(
    ...columnNames: CS
  ): { [K in CS[number]]: ColumnNamed<SN, K> } {
    const columns = {} as { [K in CS[number]]: ColumnNamed<SN, K> };
    columnNames.forEach((columnName) => {
      columns[columnName] = this.column(columnName).prepFetchFull();
    });
    return columns;
  }
  prepFetchColumnsSpecific<CS extends readonly ColumnName<SN>[]>(
    rowIndexes: number[],
    ...columnNames: CS
  ): { [K in CS[number]]: ColumnNamed<SN, K> } {
    const columns = {} as { [K in CS[number]]: ColumnNamed<SN, K> };
    columnNames.forEach((columnName) => {
      columns[columnName] =
        this.column(columnName).prepFetchSpecific(rowIndexes);
    });
    return columns;
  }
  prepFetchColumnsActive<CS extends readonly ColumnName<SN>[]>(
    ...columnNames: CS
  ): { [K in CS[number]]: ColumnNamed<SN, K> } {
    return this.prepFetchColumnsSpecific(this.rowIndexesActive, ...columnNames);
  }
  sortRowsbyColumnName(
    rows: RowNamed<SN>[],
    columnName: ColumnName<SN>,
  ): RowNamed<SN>[] {
    return rows.sort((a, b) => {
      return Arr.compareForSort(
        a.valueOrEmpty(columnName),
        b.valueOrEmpty(columnName),
      );
    });
  }
  DELETE_ALL_DATA_ROWS(): void {
    this.identified.DELETE_ALL_DATA_ROWS();
  }
  findReplace(terms: FindReplaceTerms): this {
    this.identified.findReplace(terms);
    return this;
  }
  prepFetchConditionalFormatRules(): this {
    this.identified.prepFetchConditionalFormatRules();
    return this;
  }
  conditionalFormatRules(): ConditionalFormatRule[] {
    return this.identified.conditionalFormatRules();
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
  prepFetchEditProtections(): this {
    this.identified.prepFetchEditProtections();
    return this;
  }
  editProtections(): EditProtection[] {
    return this.identified.editProtections();
  }
  addEditWarning(declaration: EditWarningDeclaration = {}): this {
    this.identified.addEditWarning(declaration);
    return this;
  }
  addEditLock(declaration: EditLockDeclaration = {}): this {
    this.identified.addEditLock(declaration);
    return this;
  }
  addEditWarningWholeSheet(
    declaration: WholeSheetEditWarningDeclaration = {},
  ): this {
    this.identified.addEditWarningWholeSheet(declaration);
    return this;
  }
  addEditLockWholeSheet(declaration: WholeSheetEditLockDeclaration = {}): this {
    this.identified.addEditLockWholeSheet(declaration);
    return this;
  }
  removeEditProtections(): this {
    this.identified.removeEditProtections();
    return this;
  }
  removeEditProtection(protection: EditProtection): this {
    this.identified.removeEditProtection(protection);
    return this;
  }
  removeEditProtectionByDescription(description: string): this {
    this.identified.removeEditProtectionByDescription(description);
    return this;
  }
  removeEditProtectionById(protectionId: number): this {
    this.identified.removeEditProtectionById(protectionId);
    return this;
  }
  anchoredA1(columnName: ColumnName<SN>): string {
    return this.column(columnName).anchoredA1();
  }
  rowsFiltered(values: Partial<SheetDataValues<SN>>): RowNamed<SN>[] {
    return this.rows.filter((row) =>
      Obj.keys(values).every(
        (columnName) => row.valueOrEmpty(columnName) === values[columnName],
      ),
    );
  }
  rowByValue<CN extends ColumnName<SN>>(
    columnName: CN,
    value: ColumnValue<SN, CN>,
  ): RowNamed<SN> {
    const rows = this.rows.filter(
      (row) => row.valueOrEmpty(columnName) === value,
    );
    if (rows.length !== 1) {
      throw new Error(
        `Expected 1 row of "${this.sheetName}" to have a "${columnName}" of "${value}", but ${rows.length} did.`,
      );
    }
    return Val.assert(rows[0], "The matching row");
  }
  appendRowWithVals(values: Partial<SheetDataValues<SN>>): RowNamed<SN> {
    const { rowIndex } = this.identified.appendRowDefault();
    return this.row(rowIndex).updateValues(values);
  }
  appendRowWithAllVals(values: SheetDataValuesAll<SN>): RowNamed<SN> {
    // Checking this subset generically costs ~70k instantiations; the Named suite pins it instead.
    return this.appendRowWithVals(
      values as unknown as Partial<SheetDataValues<SN>>,
    );
  }
  prepFetchRowIdAndName(
    this: SheetNamed<SN & SheetNameWithIdAndNameColumn>,
  ): SheetNamed<SN & SheetNameWithIdAndNameColumn> {
    this._idColumn().prepFetchFull();
    this._nameColumn().prepFetchFull();
    return this;
  }
  rowIdByName(
    this: SheetNamed<SN & SheetNameWithIdAndNameColumn>,
    name: string,
  ): RowIdByName {
    this._validateNameNotBlank(name);
    const rowIndexes = this._rowIndexesNamed(name);
    const rowIndex = rowIndexes[0];
    if (rowIndex === undefined) return { found: "none" };
    if (rowIndexes.length > 1) {
      return { found: "many", rowCount: rowIndexes.length };
    }
    return { found: "one", rowId: this._rowId(rowIndex), rowIndex };
  }
  // A blank name would match every unnamed row, which is never what a caller meant.
  private _validateNameNotBlank(name: string): void {
    if (name !== "") return;
    throw new Error(
      `Cannot look up a blank name in the name column of "${this.sheetName}".`,
    );
  }
  private _rowIndexesNamed(name: string): number[] {
    const column = this._nameColumn();
    return column.cellIndexesActive.filter(
      (rowIndex) => column.valueOrEmpty(rowIndex) === name,
    );
  }
  // A row id is bookkeeping nobody types, and the row is already in hand.
  private _rowId(rowIndex: number): string {
    const cell = this._idColumn().cell(rowIndex);
    if (cell.raw.isEmpty) {
      cell.updateToDefault();
    }
    return cell.valueNotEmpty();
  }
  // By header, since the layout values name these columns, not a column name.
  private _idColumn(): ColumnIdentified<"id"> {
    // The framework's own column, so its value type is named here.
    return new ColumnIdentified<"id">({
      ...this.identified.sheetIdentifiedProps,
      columnId: this.schema.columnIdByHeader(this.schema.idHeader),
    });
  }
  // By column id, so the name column's own value type isn't composed into the read.
  private _nameColumn(): ColumnIdentified {
    return this.identified.column(
      this.schema.columnIdByHeader(this.schema.nameHeader),
    );
  }
}
