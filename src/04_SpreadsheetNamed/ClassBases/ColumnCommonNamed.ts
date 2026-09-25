import type { ColumnName } from "../../01_SpreadsheetSchema/columnConfigsTypes";
import type { SheetName } from "../../01_SpreadsheetSchema/sheetConfigsTypes";
import { ColumnBaseNamed } from "./ColumnBaseNamed";

export abstract class ColumnCommonNamed<
  SN extends SheetName,
  CN extends ColumnName<SN>,
> extends ColumnBaseNamed<SN, CN> {
  get columnId(): string {
    return this.schema.columnId;
  }
}
