import { Obj } from "@byronbroughten/utils/obj";

import type { NotEmpty } from "../00_Source/CellValues/cellValues";
import type {
  ColumnName,
  ColumnValue,
  ColumnValueDeclared,
  SheetDataValues,
  SheetDataValuesDeclared,
  SheetDataValuesNotEmpty,
} from "../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import type { TableName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import type { Value, ValueName } from "../01_SpreadsheetSchema/configReaders/valueSchemas";
import { RowRaw } from "../02_SpreadsheetRaw/RowRaw";
import { RowIdentified } from "../03_SpreadsheetIdentified/RowIdentified";
import type { CellNamed } from "./CellNamed";
import { RowBaseNamed } from "./ClassBases/RowBaseNamed";
import { TableNamed } from "./TableNamed";

export class RowNamed<TN extends TableName> extends RowBaseNamed<TN> {
  get table(): TableNamed<TN> {
    return new TableNamed(this.sheetNamedProps);
  }
  get identified(): RowIdentified {
    return new RowIdentified({
      ...this.rowNamedProps,
      ...this.table.tableAddress,
    });
  }
  get raw(): RowRaw {
    return new RowRaw({
      ...this.table.raw.tableRawProps,
      rowIndex: this.rowIndex,
    });
  }
  cell<CN extends ColumnName<TN>>(columnName: CN): CellNamed<TN, CN> {
    return this.table.column(columnName).cell(this.rowIndex);
  }
  cellInWorking<CN extends ColumnName<TN>>(columnName: CN): boolean {
    return this.cell(columnName).inWorking;
  }
  valueOrEmpty<CN extends ColumnName<TN>>(columnName: CN): ColumnValue<TN, CN> {
    return this.cell(columnName).valueOrEmpty();
  }
  valueNotEmpty<CN extends ColumnName<TN>>(
    columnName: CN,
  ): NotEmpty<ColumnValue<TN, CN>> {
    return this.cell(columnName).valueNotEmpty();
  }
  value<CN extends ColumnName<TN>>(
    columnName: CN,
  ): ColumnValueDeclared<TN, CN> {
    return this.cell(columnName).value();
  }
  valuesOrEmpty<CN extends ColumnName<TN>>(
    columnName: CN,
    ...columnNames: CN[]
  ): SheetDataValues<TN, CN> {
    return this._valuesRefusingBlanks(
      [columnName, ...columnNames],
      () => false,
    );
  }
  valuesNotEmpty<CN extends ColumnName<TN>>(
    columnName: CN,
    ...columnNames: CN[]
  ): SheetDataValuesNotEmpty<TN, CN> {
    return this._valuesRefusingBlanks(
      [columnName, ...columnNames],
      () => true,
    ) as SheetDataValuesNotEmpty<TN, CN>;
  }
  values<CN extends ColumnName<TN>>(
    columnName: CN,
    ...columnNames: CN[]
  ): SheetDataValuesDeclared<TN, CN> {
    return this._valuesRefusingBlanks(
      [columnName, ...columnNames],
      (name) => !this.cell(name).schema.emptyValueAllowed,
    ) as SheetDataValuesDeclared<TN, CN>;
  }
  // One throw naming every refused blank, so the operator fixes the row in one pass.
  private _valuesRefusingBlanks<CN extends ColumnName<TN>>(
    columnNames: CN[],
    refusesBlank: (columnName: CN) => boolean,
  ): SheetDataValues<TN, CN> {
    const values = {} as SheetDataValues<TN, CN>;
    const refused: CN[] = [];
    columnNames.forEach((columnName) => {
      const value = this.valueOrEmpty(columnName);
      if (value === "" && refusesBlank(columnName)) {
        refused.push(columnName);
      }
      values[columnName] = value as SheetDataValues<TN, CN>[CN];
    });
    if (refused.length > 0) {
      throw new Error(this._blankColumnsMessage(refused));
    }
    return values;
  }
  private _blankColumnsMessage(columnNames: ColumnName<TN>[]): string {
    const quoted = columnNames
      .map((columnName) => `"${columnName}"`)
      .join(", ");
    const subject =
      columnNames.length === 1 ? `Column ${quoted}` : `Columns ${quoted}`;
    const verb = columnNames.length === 1 ? "is" : "are";
    return `${subject} of sheet "${this.tableName}" ${verb} empty in ${this.raw.rowLabel(this.rowIndex)}.`;
  }
  get workingCellNames(): ColumnName<TN>[] {
    return this.identified.workingColumnIds.map((columnId) =>
      this.schema.colNameByColumnId(columnId),
    );
  }
  // The column's own Empty value allowed tick is the only record of mandatoriness.
  blankRequiredColumnNames(): ColumnName<TN>[] {
    return this.workingCellNames.filter((columnName) => {
      const cell = this.cell(columnName);
      return !cell.schema.emptyValueAllowed && cell.valueOrEmpty() === "";
    });
  }
  updateToDefault(...columnNames: ColumnName<TN>[]): RowNamed<TN> {
    columnNames.forEach((columnName) =>
      this.cell(columnName).updateToDefault(),
    );
    return this;
  }
  updateCellToDefault(columnName: ColumnName<TN>): RowNamed<TN> {
    this.cell(columnName).updateToDefault();
    return this;
  }
  updateValue<CN extends ColumnName<TN>>(
    columnName: CN,
    value: ColumnValue<TN, CN>,
  ): RowNamed<TN> {
    this.cell(columnName).updateValue(value);
    return this;
  }
  get isBlank(): boolean {
    return this.identified.isBlank;
  }
  clearValues(): RowNamed<TN> {
    this.identified.clearValues();
    return this;
  }
  delete(): void {
    this.identified.delete();
  }
  setValueType<CN extends ColumnName<TN>>(
    columnName: CN,
    valueName: ValueName,
    value: Value,
  ): RowNamed<TN> {
    this.cell(columnName).setValueType(valueName, value);
    return this;
  }
  updateValues(sectionValues: Partial<SheetDataValues<TN>>): RowNamed<TN> {
    Obj.entries(sectionValues).forEach(([columnName, value]) => {
      this.updateValue(columnName, value as ColumnValue<TN, typeof columnName>);
    });
    return this;
  }
  prepFetchFull(): this {
    this.identified.prepFetchFull();
    return this;
  }
}
