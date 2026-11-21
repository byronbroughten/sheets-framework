import type { ColumnName } from "../../01_SpreadsheetSchema/configReaders/columnConfigsTypes.ts";
import type { TableName } from "../../01_SpreadsheetSchema/configReaders/tableConfigsTypes.ts";
import { ColumnBaseNamed } from "./ColumnBaseNamed.ts";

export abstract class ColumnCommonNamed<
  TN extends TableName,
  CN extends ColumnName<TN>,
> extends ColumnBaseNamed<TN, CN> {
  get columnId(): string {
    return this.schema.columnId;
  }
}
