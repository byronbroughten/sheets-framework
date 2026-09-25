import type { ColumnName } from "../../01_SpreadsheetSchema/columnConfigsTypes";
import type { SheetName } from "../../01_SpreadsheetSchema/sheetConfigsTypes";
import { ColumnBaseNamed, type ColumnNamedProps } from "./ColumnBaseNamed";

export interface CellNamedProps<
  TN extends SheetName,
  CN extends ColumnName<TN>,
> extends ColumnNamedProps<TN, CN> {
  rowIndex: number;
}

export class CellBaseNamed<
  TN extends SheetName,
  CN extends ColumnName<TN>,
> extends ColumnBaseNamed<TN, CN> {
  readonly rowIndex: number;
  constructor({ rowIndex, ...props }: CellNamedProps<TN, CN>) {
    super(props);
    this.rowIndex = rowIndex;
  }
  get cellNamedProps(): CellNamedProps<TN, CN> {
    return {
      ...this.columnNamedProps,
      rowIndex: this.rowIndex,
    };
  }
}
