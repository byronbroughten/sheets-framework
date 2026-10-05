import { Val } from "../utils/Val";
import {
  configSheetNames,
  sheetConfigsByGid,
  type TableName,
} from "./sheetConfigsTypes";
import { SheetSchema } from "./SheetSchema";
import { SpreadsheetBaseSchema } from "./SpreadsheetBaseSchema";

export class SpreadsheetSchema extends SpreadsheetBaseSchema {
  isInSheetGids(sheetGid: number): boolean {
    return sheetConfigsByGid().has(sheetGid);
  }
  get sheetNames(): TableName[] {
    return configSheetNames();
  }
  sheetByName<TN extends TableName>(sheetName: TN): SheetSchema<TN> {
    return SheetSchema.fromSheetName(sheetName);
  }
  sheetByGid(sheetGid: number): SheetSchema {
    return SheetSchema.fromSheetGid(sheetGid);
  }
  // The inverse of `ColumnSchema.fullName`; a camelCase sheet name never holds the delimiter.
  sheetByColumnFullName(fullName: string): SheetSchema {
    const sheetName = this.sheetNames.find((name) =>
      fullName.startsWith(this.combineNames(name, "")),
    );
    return SheetSchema.fromSheetName(
      Val.assert(sheetName, `sheet of column full name ${fullName}`),
    );
  }
}
