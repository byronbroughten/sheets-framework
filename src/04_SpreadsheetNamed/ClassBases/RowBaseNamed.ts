import type { TableName } from "../../01_SpreadsheetSchema/sheetConfigsTypes";
import { SheetSchema } from "../../01_SpreadsheetSchema/SheetSchema";
import { SheetBaseNamed, type SheetNamedProps } from "./SheetBaseNamed";

export interface RowNamedProps<
  TN extends TableName,
> extends SheetNamedProps<TN> {
  rowIndex: number;
}

export class RowBaseNamed<TN extends TableName> extends SheetBaseNamed<TN> {
  readonly rowIndex: number;
  constructor({ rowIndex, ...props }: RowNamedProps<TN>) {
    super(props);
    this.rowIndex = rowIndex;
  }
  get schema(): SheetSchema<TN> {
    return SheetSchema.fromSheetName(this.sheetName);
  }
  get rowNamedProps(): RowNamedProps<TN> {
    return {
      ...this.sheetNamedProps,
      rowIndex: this.rowIndex,
    };
  }
}
