import type { ColumnName } from "../../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import type { TableName } from "../../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { ColumnBaseNamed, type ColumnNamedProps } from "./ColumnBaseNamed";

export interface CellNamedProps<
  TN extends TableName,
  CN extends ColumnName<TN>,
> extends ColumnNamedProps<TN, CN> {
  rowIndex: number;
}

export class CellBaseNamed<
  TN extends TableName,
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
