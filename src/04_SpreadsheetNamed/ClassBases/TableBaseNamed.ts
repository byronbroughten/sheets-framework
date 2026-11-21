import type { TableName } from "../../01_SpreadsheetSchema/configReaders/tableConfigsTypes.ts";
import {
  SpreadsheetBaseNamed,
  type SpreadsheetNamedProps,
} from "./SpreadsheetBaseNamed.ts";

export interface TableNamedProps<
  TN extends TableName,
> extends SpreadsheetNamedProps {
  tableName: TN;
}

export class TableBaseNamed<TN extends TableName> extends SpreadsheetBaseNamed {
  declare readonly tableName: TN;
  constructor({ tableName, ...props }: TableNamedProps<TN>) {
    super(props);
    this.tableName = tableName;
  }
  get sheetNamedProps(): TableNamedProps<TN> {
    return {
      tableName: this.tableName,
      ...this.spreadsheetNamedProps,
    };
  }
}
