import type { TableName } from "../../01_SpreadsheetSchema/sheetConfigsTypes";
import {
  SpreadsheetBaseNamed,
  type SpreadsheetNamedProps,
} from "./SpreadsheetBaseNamed";

export interface SheetNamedProps<
  TN extends TableName,
> extends SpreadsheetNamedProps {
  sheetName: TN;
}

export class SheetBaseNamed<TN extends TableName> extends SpreadsheetBaseNamed {
  readonly sheetName: TN;
  constructor({ sheetName, ...props }: SheetNamedProps<TN>) {
    super(props);
    this.sheetName = sheetName;
  }
  get sheetNamedProps(): SheetNamedProps<TN> {
    return {
      sheetName: this.sheetName,
      ...this.spreadsheetNamedProps,
    };
  }
}
