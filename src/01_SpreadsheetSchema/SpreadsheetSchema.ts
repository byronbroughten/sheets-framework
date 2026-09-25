import {
  configSheetNames,
  sheetConfigsByGid,
  type SheetName,
} from "./sheetConfigsTypes";
import { SheetSchema } from "./SheetSchema";
import { SpreadsheetBaseSchema } from "./SpreadsheetBaseSchema";

export class SpreadsheetSchema extends SpreadsheetBaseSchema {
  isInSheetGids(sheetGid: number): boolean {
    return sheetConfigsByGid().has(sheetGid);
  }
  get sheetNames(): SheetName[] {
    return configSheetNames();
  }
  sheetByName<SN extends SheetName>(sheetName: SN): SheetSchema<SN> {
    return SheetSchema.fromSheetName(sheetName);
  }
  sheetByGid(sheetGid: number): SheetSchema {
    return SheetSchema.fromSheetGid(sheetGid);
  }
}
