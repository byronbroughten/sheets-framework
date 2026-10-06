import type { ColumnName } from "../01_SpreadsheetSchema/columnConfigsTypes";
import type { TableNameSimple } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import { TableSchema } from "../01_SpreadsheetSchema/TableSchema";
import { TableBaseNamed } from "../04_SpreadsheetNamed/ClassBases/TableBaseNamed";
import type { ColumnNamed } from "../04_SpreadsheetNamed/ColumnNamed";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import type { TableNamed } from "../04_SpreadsheetNamed/TableNamed";
import type { ConfigSyncState, OperatorProps } from "./SpreadsheetBaseOperator";

export interface SheetOperatorProps<
  TN extends TableNameSimple,
> extends OperatorProps {
  tableName: TN;
}

export class GenericTableOperator<
  TN extends TableNameSimple,
> extends TableBaseNamed<TN> {
  protected configSyncState: ConfigSyncState;
  constructor({ configSyncState, ...rest }: SheetOperatorProps<TN>) {
    super(rest);
    this.configSyncState = configSyncState;
  }
  get operatorProps(): OperatorProps {
    return {
      ...this.spreadsheetNamedProps,
      configSyncState: this.configSyncState,
    };
  }
  get ss(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  get table(): TableNamed<TN> {
    return this.ss.table(this.tableName);
  }
  get schema(): TableSchema<TN> {
    return TableSchema.fromSheetName(this.tableName);
  }
  column<CN extends ColumnName<TN>>(columnName: CN): ColumnNamed<TN, CN> {
    return this.table.column(columnName);
  }
}
