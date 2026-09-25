import type { ColumnName } from "../../01_SpreadsheetSchema/columnConfigsTypes";
import { ColumnSchema } from "../../01_SpreadsheetSchema/ColumnSchema";
import type { SheetName } from "../../01_SpreadsheetSchema/sheetConfigsTypes";
import { SheetBaseNamed, type SheetNamedProps } from "./SheetBaseNamed";

export interface ColumnNamedProps<
  TN extends SheetName,
  CN extends ColumnName<TN>,
> extends SheetNamedProps<TN> {
  columnName: CN;
}

export class ColumnBaseNamed<
  TN extends SheetName,
  CN extends ColumnName<TN>,
> extends SheetBaseNamed<TN> {
  readonly columnName: CN;
  constructor(props: ColumnNamedProps<TN, CN>) {
    super(props);
    this.columnName = props.columnName;
  }
  get schema(): ColumnSchema<TN, CN> {
    return ColumnSchema.fromColumnName(this.sheetName, this.columnName);
  }
  get columnNamedProps(): ColumnNamedProps<TN, CN> {
    return {
      ...this.sheetNamedProps,
      columnName: this.columnName,
    };
  }
}
