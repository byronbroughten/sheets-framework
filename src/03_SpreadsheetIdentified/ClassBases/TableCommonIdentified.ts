import type { TableName } from "../../01_SpreadsheetSchema/sheetConfigsTypes";
import { SheetSchema } from "../../01_SpreadsheetSchema/SheetSchema";
import { TableBaseIdentified } from "./TableBaseIdentified";

export abstract class TableCommonIdentified extends TableBaseIdentified {
  get schema(): SheetSchema {
    return SheetSchema.fromSheetGid(this.sheetGid);
  }
  get sheetName(): TableName {
    return this.schema.sheetName;
  }
}
