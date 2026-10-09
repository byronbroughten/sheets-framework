import type { ColumnName } from "../../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import type { TableName } from "../../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { ColumnBaseNamed } from "./ColumnBaseNamed";

export abstract class ColumnCommonNamed<
  TN extends TableName,
  CN extends ColumnName<TN>,
> extends ColumnBaseNamed<TN, CN> {
  get columnId(): string {
    return this.schema.columnId;
  }
}
