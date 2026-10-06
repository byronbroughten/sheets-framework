import type { NotEmpty } from "../00_Source/CellValues/cellValues";
import type {
  ColumnName,
  ColumnValue,
  ColumnValueDeclared,
  SheetDataValues,
} from "../01_SpreadsheetSchema/columnConfigsTypes";
import type { TableName } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import type { Value, ValueName } from "../01_SpreadsheetSchema/valueSchemas";
import { RowRaw } from "../02_SpreadsheetRaw/RowRaw";
import { RowIdentified } from "../03_SpreadsheetIdentified/RowIdentified";
import { Obj } from "../utils/Obj";
import type { CellNamed } from "./CellNamed";
import { RowBaseNamed } from "./ClassBases/RowBaseNamed";
import { TableNamed } from "./TableNamed";

export class RowNamed<TN extends TableName> extends RowBaseNamed<TN> {
  get sheet(): TableNamed<TN> {
    return new TableNamed(this.sheetNamedProps);
  }
  get identified(): RowIdentified {
    return new RowIdentified({
      ...this.rowNamedProps,
      sheetGid: this.sheet.sheetGid,
    });
  }
  get raw(): RowRaw {
    return new RowRaw({
      ...this.sheet.raw.tableRawProps,
      rowIndex: this.rowIndex,
    });
  }
  cell<CN extends ColumnName<TN>>(columnName: CN): CellNamed<TN, CN> {
    return this.sheet.column(columnName).cell(this.rowIndex);
  }
  cellIsActive<CN extends ColumnName<TN>>(columnName: CN): boolean {
    return this.cell(columnName).isActive;
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
  valuesOrEmpty<CN extends ColumnName<TN> = ColumnName<TN>>(
    ...columnNames: readonly CN[]
  ): SheetDataValues<TN, CN> {
    const keys =
      columnNames.length > 0 ? columnNames : (this.activeCellNames as CN[]);
    return keys.reduce(
      (values, columnName) => {
        (values[columnName] as SheetDataValues<TN, CN>[CN]) = this.valueOrEmpty(
          columnName,
        ) as SheetDataValues<TN, CN>[CN];
        return values;
      },
      {} as SheetDataValues<TN, CN>,
    );
  }
  get activeCellNames(): ColumnName<TN>[] {
    return this.identified.activeColumnIds.map((columnId) =>
      this.schema.colNameByColumnId(columnId),
    );
  }
  // The column's own Empty value allowed tick is the only record of mandatoriness.
  blankRequiredColumnNames(): ColumnName<TN>[] {
    return this.activeCellNames.filter((columnName) => {
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
