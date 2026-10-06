import type { TableName } from "../../01_SpreadsheetSchema/sheetConfigsTypes";
import {
  SpreadsheetBaseNamed,
  type SpreadsheetNamedProps,
} from "./SpreadsheetBaseNamed";

export interface TableNamedProps<
  TN extends TableName,
> extends SpreadsheetNamedProps {
  sheetName: TN;
}

export class TableBaseNamed<TN extends TableName> extends SpreadsheetBaseNamed {
  readonly sheetName: TN;
  constructor({ sheetName, ...props }: TableNamedProps<TN>) {
    super(props);
    this.sheetName = sheetName;
  }
  get sheetNamedProps(): TableNamedProps<TN> {
    return {
      sheetName: this.sheetName,
      ...this.spreadsheetNamedProps,
    };
  }
}
