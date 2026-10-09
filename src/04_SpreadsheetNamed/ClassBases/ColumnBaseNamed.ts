import type { ColumnName } from "../../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import { ColumnSchema } from "../../01_SpreadsheetSchema/configReaders/ColumnSchema";
import type { TableName } from "../../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { TableBaseNamed, type TableNamedProps } from "./TableBaseNamed";

export interface ColumnNamedProps<
  TN extends TableName,
  CN extends ColumnName<TN>,
> extends TableNamedProps<TN> {
  columnName: CN;
}

export class ColumnBaseNamed<
  TN extends TableName,
  CN extends ColumnName<TN>,
> extends TableBaseNamed<TN> {
  readonly columnName: CN;
  constructor(props: ColumnNamedProps<TN, CN>) {
    super(props);
    this.columnName = props.columnName;
  }
  get schema(): ColumnSchema<TN, CN> {
    return ColumnSchema.fromColumnName(this.tableName, this.columnName);
  }
  get columnNamedProps(): ColumnNamedProps<TN, CN> {
    return {
      ...this.sheetNamedProps,
      columnName: this.columnName,
    };
  }
}
