import { Val } from "../utils/Val";
import {
  configSheetNames,
  sheetConfigsByGid,
  type TableName,
} from "./sheetConfigsTypes";
import { SpreadsheetBaseSchema } from "./SpreadsheetBaseSchema";
import { TableSchema } from "./TableSchema";

export class SpreadsheetSchema extends SpreadsheetBaseSchema {
  isInSheetGids(sheetGid: number): boolean {
    return sheetConfigsByGid().has(sheetGid);
  }
  get sheetNames(): TableName[] {
    return configSheetNames();
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
