import { Val } from "../utils/Val";
import { SpreadsheetBaseSchema } from "./SpreadsheetBaseSchema";
import {
  configTableNames,
  tableConfigsByGid,
  type TableName,
} from "./tableConfigsTypes";
import { TableSchema } from "./TableSchema";

export class SpreadsheetSchema extends SpreadsheetBaseSchema {
  isInSheetGids(sheetGid: number): boolean {
    return tableConfigsByGid().has(sheetGid);
  }
  get sheetNames(): TableName[] {
    return configTableNames();
  }
  sheetByName<TN extends TableName>(sheetName: TN): TableSchema<TN> {
    return TableSchema.fromSheetName(sheetName);
  }
  sheetByGid(sheetGid: number): TableSchema {
    return TableSchema.fromSheetGid(sheetGid);
  }
  // The inverse of `ColumnSchema.fullName`; a camelCase sheet name never holds the delimiter.
  sheetByColumnFullName(fullName: string): TableSchema {
    const sheetName = this.sheetNames.find((name) =>
      fullName.startsWith(this.combineNames(name, "")),
    );
    return TableSchema.fromSheetName(
      Val.assert(sheetName, `sheet of column full name ${fullName}`),
    );
  }
}
