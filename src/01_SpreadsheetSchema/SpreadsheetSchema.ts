import type { SheetEdit } from "../00_Source/PlatformEvents/sheetEdit";
import { Val } from "../utils/Val";
import { SpreadsheetBaseSchema } from "./SpreadsheetBaseSchema";
import {
  configTableNames,
  tableConfigsByGid,
  tableKeysByGid,
  type TableName,
} from "./tableConfigsTypes";
import { TableOrigin } from "./TableOrigin";
import { TableSchema } from "./TableSchema";

export class SpreadsheetSchema extends SpreadsheetBaseSchema {
  isInSheetGids(sheetGid: number): boolean {
    return tableConfigsByGid().has(sheetGid);
  }
  get sheetNames(): TableName[] {
    return configTableNames();
  }
  sheetByName<TN extends TableName>(tableName: TN): TableSchema<TN> {
    return TableSchema.fromSheetName(tableName);
  }
  sheetByGid(sheetGid: number): TableSchema {
    return TableSchema.fromSheetGid(sheetGid);
  }
  tablesOnGid(sheetGid: number): TableSchema[] {
    return (tableKeysByGid().get(sheetGid) ?? []).map((tableKey) =>
      TableSchema.fromSheetName(tableKey),
    );
  }
  // A sheet the configs don't record has only the spot the framework creates Tables at.
  presumedOrigin(sheetGid: number): TableOrigin {
    if (!this.isInSheetGids(sheetGid)) return TableOrigin.expected();
    return this.sheetByGid(sheetGid).recordedOrigin;
  }
  // From recorded positions alone, so the edit trigger answers without a fetch.
  tableWithActionCell({
    sheetGid,
    rowIndexBase0,
    colIndexBase0,
  }: SheetEdit): TableSchema | undefined {
    return this.tablesOnGid(sheetGid).find((table) =>
      table.holdsActionCellAt(rowIndexBase0, colIndexBase0),
    );
  }
  // The inverse of `ColumnSchema.fullName`; a camelCase sheet name never holds the delimiter.
  sheetByColumnFullName(fullName: string): TableSchema {
    const tableName = this.sheetNames.find((name) =>
      fullName.startsWith(this.combineNames(name, "")),
    );
    return TableSchema.fromSheetName(
      Val.assert(tableName, `sheet of column full name ${fullName}`),
    );
  }
}
