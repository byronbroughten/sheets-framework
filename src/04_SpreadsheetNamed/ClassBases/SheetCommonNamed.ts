import type { TableName } from "../../01_SpreadsheetSchema/sheetConfigsTypes";
import { TableSchema } from "../../01_SpreadsheetSchema/TableSchema";
import { SheetBaseNamed } from "./SheetBaseNamed";

export abstract class SheetCommonNamed<
  TN extends TableName,
> extends SheetBaseNamed<TN> {
  get schema(): TableSchema<TN> {
    return TableSchema.fromSheetName(this.sheetName);
  }
  get sheetGid(): number {
    return this.schema.sheetGid;
  }
}
