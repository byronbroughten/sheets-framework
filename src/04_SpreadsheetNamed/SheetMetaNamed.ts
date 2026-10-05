import type { UniformRowName } from "../00_Source/CellValues/cellValues";
import type { ColumnName } from "../01_SpreadsheetSchema/columnConfigsTypes";
import type { TableName } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import type { SheetMetaRaw } from "../02_SpreadsheetRaw/SheetMetaRaw";
import { SheetMetaIdentified } from "../03_SpreadsheetIdentified/SheetMetaIdentified";
import type { UniformRowIdentified } from "../03_SpreadsheetIdentified/UniformRowIdentified";
import { SheetCommonNamed } from "./ClassBases/SheetCommonNamed";
import { ColumnMetaNamed } from "./ColumnMetaNamed";
import { SheetNamed } from "./SheetNamed";
import { SpreadsheetNamed } from "./SpreadsheetNamed";

export class SheetMetaNamed<
  TN extends TableName = TableName,
> extends SheetCommonNamed<TN> {
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
  get primary(): SheetNamed<TN> {
    return new SheetNamed(this.sheetNamedProps);
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
  uniformRow<UN extends UniformRowName>(rowName: UN): UniformRowIdentified<UN> {
    return this.identified.uniformRow(rowName);
  }
  isActiveColumnId(columnId: string): boolean {
    return this.identified.isActiveColumnId(columnId);
  }
  addMissingColumnIds(): number {
    return this.identified.addMissingColumnIds();
  }
}
