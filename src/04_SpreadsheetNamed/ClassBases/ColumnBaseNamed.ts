import type { ColumnName } from "../../01_SpreadsheetSchema/configReaders/columnConfigsTypes.ts";
import { ColumnSchema } from "../../01_SpreadsheetSchema/configReaders/ColumnSchema.ts";
import type { TableName } from "../../01_SpreadsheetSchema/configReaders/tableConfigsTypes.ts";
import { TableBaseNamed, type TableNamedProps } from "./TableBaseNamed.ts";

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
  declare readonly columnName: CN;
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
