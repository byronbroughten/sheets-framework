import type { SheetName } from "../../01_SpreadsheetSchema/sheetConfigsTypes";
import { SheetSchema } from "../../01_SpreadsheetSchema/SheetSchema";
import { SheetBaseIdentified } from "./SheetBaseIdentified";

export abstract class SheetCommonIdentified extends SheetBaseIdentified {
  get schema(): SheetSchema {
    return SheetSchema.fromSheetGid(this.sheetGid);
  }
  get sheetName(): SheetName {
    return this.schema.sheetName;
  }
}
