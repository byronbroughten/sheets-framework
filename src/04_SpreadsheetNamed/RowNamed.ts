import type { NotEmpty } from "../00_Source/CellValues/cellValues";
import type {
  ColumnName,
  ColumnValue,
  ColumnValueDeclared,
  SheetDataValues,
} from "../01_SpreadsheetSchema/columnConfigsTypes";
import type { SheetName } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import type { Value, ValueName } from "../01_SpreadsheetSchema/valueSchemas";
import { RowRaw } from "../02_SpreadsheetRaw/RowRaw";
import { RowIdentified } from "../03_SpreadsheetIdentified/RowIdentified";
import { Obj } from "../utils/Obj";
import type { CellNamed } from "./CellNamed";
import { RowBaseNamed } from "./ClassBases/RowBaseNamed";
import { SheetNamed } from "./SheetNamed";

export class RowNamed<SN extends SheetName> extends RowBaseNamed<SN> {
  get sheet(): SheetNamed<SN> {
    return new SheetNamed(this.sheetNamedProps);
  }
  get identified(): RowIdentified {
    return new RowIdentified({
      ...this.rowNamedProps,
      sheetGid: this.sheet.sheetGid,
    });
  }
  get raw(): RowRaw {
    return new RowRaw({
      ...this.sheet.raw.sheetRawProps,
      rowIndex: this.rowIndex,
    });
  }
  cell<CN extends ColumnName<SN>>(columnName: CN): CellNamed<SN, CN> {
    return this.sheet.column(columnName).cell(this.rowIndex);
  }
  cellIsActive<CN extends ColumnName<SN>>(columnName: CN): boolean {
    return this.cell(columnName).isActive;
  }
  valueOrEmpty<CN extends ColumnName<SN>>(columnName: CN): ColumnValue<SN, CN> {
    return this.cell(columnName).valueOrEmpty();
  }
  valueNotEmpty<CN extends ColumnName<SN>>(
    columnName: CN,
  ): NotEmpty<ColumnValue<SN, CN>> {
    return this.cell(columnName).valueNotEmpty();
  }
  value<CN extends ColumnName<SN>>(
    columnName: CN,
  ): ColumnValueDeclared<SN, CN> {
    return this.cell(columnName).value();
  }
  valuesOrEmpty<CN extends ColumnName<SN> = ColumnName<SN>>(
    ...columnNames: readonly CN[]
  ): SheetDataValues<SN, CN> {
    const keys =
      columnNames.length > 0 ? columnNames : (this.activeCellNames as CN[]);
    return keys.reduce(
      (values, columnName) => {
        (values[columnName] as SheetDataValues<SN, CN>[CN]) = this.valueOrEmpty(
          columnName,
        ) as SheetDataValues<SN, CN>[CN];
        return values;
      },
      {} as SheetDataValues<SN, CN>,
    );
  }
  get activeCellNames(): ColumnName<SN>[] {
    return this.identified.activeColumnIds.map((columnId) =>
      this.schema.colNameByColumnId(columnId),
    );
  }
  // The column's own Empty value allowed tick is the only record of mandatoriness.
  blankRequiredColumnNames(): ColumnName<SN>[] {
    return this.activeCellNames.filter((columnName) => {
      const cell = this.cell(columnName);
      return !cell.schema.emptyValueAllowed && cell.valueOrEmpty() === "";
    });
  }
  updateToDefault(...columnNames: ColumnName<SN>[]): RowNamed<SN> {
    columnNames.forEach((columnName) =>
      this.cell(columnName).updateToDefault(),
    );
    return this;
  }
  updateCellToDefault(columnName: ColumnName<SN>): RowNamed<SN> {
    this.cell(columnName).updateToDefault();
    return this;
  }
  updateValue<CN extends ColumnName<SN>>(
    columnName: CN,
    value: ColumnValue<SN, CN>,
  ): RowNamed<SN> {
    this.cell(columnName).updateValue(value);
    return this;
  }
  get isBlank(): boolean {
    return this.identified.isBlank;
  }
  clearValues(): RowNamed<SN> {
    this.identified.clearValues();
    return this;
  }
  delete(): void {
    this.identified.delete();
  }
  setValueType<CN extends ColumnName<SN>>(
    columnName: CN,
    valueName: ValueName,
    value: Value,
  ): RowNamed<SN> {
    this.cell(columnName).setValueType(valueName, value);
    return this;
  }
  updateValues(sectionValues: Partial<SheetDataValues<SN>>): RowNamed<SN> {
    Obj.entries(sectionValues).forEach(([columnName, value]) => {
      this.updateValue(columnName, value as ColumnValue<SN, typeof columnName>);
    });
    return this;
  }
  prepFetchFull(): this {
    this.identified.prepFetchFull();
    return this;
  }
}
