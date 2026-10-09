import { Val } from "@byronbroughten/utils/val";

import type { SheetEdit } from "../../00_Source/PlatformEvents/sheetEdit";
import { headerZone } from "../headerZone";
import { SpreadsheetBaseSchema } from "../SpreadsheetBaseSchema";
import {
  configTableNames,
  tableKeysByGid,
  type TableName,
} from "./tableConfigsTypes";
import { TableSchema } from "./TableSchema";

export class SpreadsheetSchema extends SpreadsheetBaseSchema {
  isInSheetGids(sheetGid: number): boolean {
    return tableKeysByGid().has(sheetGid);
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
  // Absent when the configs record no Table or several on that sheet.
  loneTableOnGid(sheetGid: number): TableSchema | undefined {
    const [table, ...otherTables] = this.tablesOnGid(sheetGid);
    return otherTables.length === 0 ? table : undefined;
  }
  // Without a fetch, so the edit trigger's pre-check is free; the dispatch confirms it on the live Table.
  mayHoldActionCell({ sheetGid, rowIndexBase0 }: SheetEdit): boolean {
    return (
      this.isInSheetGids(sheetGid) && headerZone.holdsActionRow(rowIndexBase0)
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
