import type { SheetName } from "../../01_SpreadsheetSchema/sheetConfigsTypes";
import { SheetSchema } from "../../01_SpreadsheetSchema/SheetSchema";
import { SheetBaseNamed } from "./SheetBaseNamed";

export abstract class SheetCommonNamed<
  SN extends SheetName,
> extends SheetBaseNamed<SN> {
  get schema(): SheetSchema<SN> {
    return SheetSchema.fromSheetName(this.sheetName);
  }
  get sheetGid(): number {
    return this.schema.sheetGid;
  }
}
