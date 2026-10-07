import type { ColumnName } from "../01_SpreadsheetSchema/columnConfigsTypes";
import type { TableName } from "../01_SpreadsheetSchema/tableConfigsTypes";
import type { SheetMetaRaw } from "../02_SpreadsheetRaw/SheetMetaRaw";
import { SheetMetaIdentified } from "../03_SpreadsheetIdentified/SheetMetaIdentified";
import { TableCommonNamed } from "./ClassBases/TableCommonNamed";
import { ColumnMetaNamed } from "./ColumnMetaNamed";
import { SpreadsheetNamed } from "./SpreadsheetNamed";
import { TableNamed } from "./TableNamed";

export class SheetMetaNamed<
  TN extends TableName = TableName,
> extends TableCommonNamed<TN> {
  get spreadsheet(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  get raw(): SheetMetaRaw {
    return this.spreadsheet.raw.sheetMeta(this.sheetGid);
  }
  get identified(): SheetMetaIdentified {
    return new SheetMetaIdentified({
      ...this.sheetNamedProps,
      sheetGid: this.sheetGid,
    });
  }
  get primary(): TableNamed<TN> {
    return new TableNamed(this.sheetNamedProps);
  }
  get activeColumnIds(): string[] {
    return this.raw.activeColumnIds;
  }
  column<CN extends ColumnName<TN>>(columnName: CN): ColumnMetaNamed<TN, CN> {
    return new ColumnMetaNamed({
      ...this.sheetNamedProps,
      columnName,
    });
  }
  columnByIndex(colIndex: number): ColumnMetaNamed<TN> {
    const columnId = this.identified.columnIdByIndex(colIndex);
    const columnName = this.schema.colNameByColumnId(columnId);
    return this.column(columnName);
  }
  isActiveColumnId(columnId: string): boolean {
    return this.identified.isActiveColumnId(columnId);
  }
  addMissingColumnIds(): number {
    return this.identified.addMissingColumnIds();
  }
}
