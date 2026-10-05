import type { TableName } from "../../01_SpreadsheetSchema/sheetConfigsTypes";
import { SheetSchema } from "../../01_SpreadsheetSchema/SheetSchema";
import { SheetBaseNamed } from "./SheetBaseNamed";

export abstract class SheetCommonNamed<
  TN extends TableName,
> extends SheetBaseNamed<TN> {
  get schema(): SheetSchema<TN> {
    return SheetSchema.fromSheetName(this.sheetName);
  }
  get sheetGid(): number {
    return this.schema.sheetGid;
  }
}
