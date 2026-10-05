import type { TableName } from "../../01_SpreadsheetSchema/sheetConfigsTypes";
import { TableSchema } from "../../01_SpreadsheetSchema/TableSchema";
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
  get schema(): TableSchema<TN> {
    return TableSchema.fromSheetName(this.sheetName);
  }
  get rowNamedProps(): RowNamedProps<TN> {
    return {
      ...this.sheetNamedProps,
      rowIndex: this.rowIndex,
    };
  }
}
