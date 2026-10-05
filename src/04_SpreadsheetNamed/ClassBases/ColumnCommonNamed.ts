import type { ColumnName } from "../../01_SpreadsheetSchema/columnConfigsTypes";
import type { TableName } from "../../01_SpreadsheetSchema/sheetConfigsTypes";
import { ColumnBaseNamed } from "./ColumnBaseNamed";

export abstract class ColumnCommonNamed<
  TN extends TableName,
  CN extends ColumnName<TN>,
> extends ColumnBaseNamed<TN, CN> {
  get columnId(): string {
    return this.schema.columnId;
  }
}
