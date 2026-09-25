import { GenericSheetOperator } from "./GenericSheetOperator";
import {
  type OperatorProps,
  SpreadsheetBaseOperator,
} from "./SpreadsheetBaseOperator";

export class SpreadsheetConfigOperator extends GenericSheetOperator<"spreadsheetConfig"> {
  constructor(props: OperatorProps) {
    super({
      sheetName: "spreadsheetConfig",
      ...props,
    });
  }
  static init(): SpreadsheetConfigOperator {
    return new SpreadsheetConfigOperator(
      SpreadsheetBaseOperator.initOperatorProps(),
    );
  }
  validateExactlyOneDataRow(): void {
    const dataRowCount = this.sheet.raw.dataRowCountAfterFlush;
    if (dataRowCount !== 1) {
      throw new Error(
        `Spreadsheet Config Table must have exactly one data row; found ${dataRowCount}.`,
      );
    }
  }
}
