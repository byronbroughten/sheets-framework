import { Arr } from "@byronbroughten/utils/arr";
import { Obj } from "@byronbroughten/utils/obj";
import { Val } from "@byronbroughten/utils/val";

import type { ConditionalFormatDeclaration } from "../00_Source/RawSource/ConditionalFormat";
import type {
  EditLockDeclaration,
  EditWarningDeclaration,
} from "../00_Source/RawSource/EditProtection";
import type { FindReplaceTerms } from "../00_Source/RawSource/RawSource";
import type {
  ColumnName,
  ColumnValue,
  SheetDataValues,
  SheetDataValuesAll,
} from "../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import type { TableName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import type { HeadRole } from "../01_SpreadsheetSchema/headRows";
import type { TableRaw } from "../02_SpreadsheetRaw/TableRaw";
import { ColumnIdentified } from "../03_SpreadsheetIdentified/ColumnIdentified";
import type { HeadRowIdentified } from "../03_SpreadsheetIdentified/HeadRowIdentified";
import { TableIdentified } from "../03_SpreadsheetIdentified/TableIdentified";
import { TableCommonNamed } from "./ClassBases/TableCommonNamed";
import { ColumnNamed } from "./ColumnNamed";
import { RowNamed } from "./RowNamed";
import { SheetNamed } from "./SheetNamed";
import type { SheetNameWithIdAndNameColumn } from "./TableNameGroups";
import type { RowIdByName } from "./Types/RowIdByName";

/**
 * Name-addressed Table: data rows, append, named columns.
 * Head rows are headRow(role); descriptive facts are the Raw-only
 * raw.profile; column IDs are identified.ensureColumnIdsAreFetched and
 * identified.addMissingColumnIds.
 * docs/architecture/class-chains.md
 */
export class TableNamed<
  TN extends TableName = TableName,
> extends TableCommonNamed<TN> {
  get raw(): TableRaw {
    return this.identified.raw;
  }
  get identified(): TableIdentified {
    return new TableIdentified({
      ...this.sheetNamedProps,
      ...this.tableAddress,
    });
  }
  get sheet(): SheetNamed {
    return new SheetNamed({
      ...this.spreadsheetNamedProps,
      sheetGid: this.sheetGid,
    });
  }
  get workingRowIndexes(): number[] {
    return this.identified.workingRowIndexes;
  }
  get workingRowIndexesWithData(): number[] {
    return this.identified.workingRowIndexesWithData;
  }
  get rowIndexesFullWithData(): number[] {
    return this.identified.rowIndexesFullWithData;
  }
  get rows(): RowNamed<TN>[] {
    return this.identified.rows.map((row) => this.row(row.rowIndex));
  }
  get topRow(): RowNamed<TN> {
    return this.row(0);
  }
  row(rowIndex: number): RowNamed<TN> {
    return new RowNamed({
      ...this.sheetNamedProps,
      rowIndex,
    });
  }
  headRow<HR extends HeadRole>(headRole: HR): HeadRowIdentified<HR> {
    return this.identified.headRow(headRole);
  }
  column<CN extends ColumnName<TN>>(columnName: CN): ColumnNamed<TN, CN> {
    return new ColumnNamed({
      ...this.sheetNamedProps,
      columnName,
    });
  }
  // By id, so the column name's value type isn't composed into the result.
  columnIdentified(columnName: ColumnName<TN>): ColumnIdentified {
    const { columnId } = this.schema.columnByName(columnName);
    return this.identified.column(columnId);
  }
  columns<CS extends readonly ColumnName<TN>[]>(
    ...columnNames: CS
  ): { [K in CS[number]]: ColumnNamed<TN, K> } {
    const columns = {} as { [K in CS[number]]: ColumnNamed<TN, K> };
    columnNames.forEach((columnName) => {
      columns[columnName] = this.column(columnName);
    });
    return columns;
  }
  prepFetchColumnsFull<CS extends readonly ColumnName<TN>[]>(
    ...columnNames: CS
  ): { [K in CS[number]]: ColumnNamed<TN, K> } {
    return this._columnsPrepped(columnNames, (column) =>
      column.prepFetchFull(),
    );
  }
  prepFetchColumnsSpecific<CS extends readonly ColumnName<TN>[]>(
    rowIndexes: number[],
    ...columnNames: CS
  ): { [K in CS[number]]: ColumnNamed<TN, K> } {
    return this._columnsPrepped(columnNames, (column) =>
      column.prepFetchSpecific(rowIndexes),
    );
  }
  private _columnsPrepped<CS extends readonly ColumnName<TN>[]>(
    columnNames: CS,
    prep: (column: ColumnNamed<TN, CS[number]>) => void,
  ): { [K in CS[number]]: ColumnNamed<TN, K> } {
    const columns = this.columns(...columnNames);
    Obj.values(columns).forEach(prep);
    return columns;
  }
  sortRowsByColumnName(
    rows: RowNamed<TN>[],
    columnName: ColumnName<TN>,
  ): RowNamed<TN>[] {
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
  addConditionalFormatRule(declaration: ConditionalFormatDeclaration): this {
    this.identified.addConditionalFormatRule(declaration);
    return this;
  }
  removeConditionalFormatRules(): this {
    this.identified.removeConditionalFormatRules();
    return this;
  }
  addEditWarning(declaration: EditWarningDeclaration = {}): this {
    this.identified.addEditWarning(declaration);
    return this;
  }
  addEditLock(declaration: EditLockDeclaration = {}): this {
    this.identified.addEditLock(declaration);
    return this;
  }
  removeEditProtections(): this {
    this.identified.removeEditProtections();
    return this;
  }
  rowsFiltered(values: Partial<SheetDataValues<TN>>): RowNamed<TN>[] {
    return this.rows.filter((row) =>
      Obj.keys(values).every(
        (columnName) => row.valueOrEmpty(columnName) === values[columnName],
      ),
    );
  }
  rowByValue<CN extends ColumnName<TN>>(
    columnName: CN,
    value: ColumnValue<TN, CN>,
  ): RowNamed<TN> {
    const values: Partial<SheetDataValues<TN>> = {};
    values[columnName] = value;
    const rows = this.rowsFiltered(values);
    if (rows.length !== 1) {
      throw new Error(
        `Expected 1 row of "${this.tableName}" to have a "${columnName}" of "${value}", but ${rows.length} did.`,
      );
    }
    return Val.assert(rows[0], "The matching row");
  }
  appendRowWithVals(values: Partial<SheetDataValues<TN>>): RowNamed<TN> {
    const { rowIndex } = this.identified.appendRowDefault();
    return this.row(rowIndex).updateValues(values);
  }
  appendRowWithAllVals(values: SheetDataValuesAll<TN>): RowNamed<TN> {
    // Checking this subset generically costs ~70k instantiations; the Named suite pins it instead.
    return this.appendRowWithVals(
      values as unknown as Partial<SheetDataValues<TN>>,
    );
  }
  prepFetchRowIdAndName(
    this: TableNamed<TN & SheetNameWithIdAndNameColumn>,
  ): TableNamed<TN & SheetNameWithIdAndNameColumn> {
    this._idColumn().prepFetchFull();
    this._nameColumn().prepFetchFull();
    return this;
  }
  rowIdByName(
    this: TableNamed<TN & SheetNameWithIdAndNameColumn>,
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
      `Cannot look up a blank name in the name column of "${this.tableName}".`,
    );
  }
  private _rowIndexesNamed(name: string): number[] {
    const column = this._nameColumn();
    return column.workingCellIndexes.filter(
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
  // By header, since tableLayout names these columns, not a column name.
  private _idColumn(): ColumnIdentified<"id"> {
    // The framework's own column, so its value type is named here.
    return new ColumnIdentified<"id">({
      ...this.identified.tableIdentifiedProps,
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
