import type { ColumnName } from "../01_SpreadsheetSchema/columnConfigsTypes";
import type { TableNameSimple } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import { SheetSchema } from "../01_SpreadsheetSchema/SheetSchema";
import { SheetBaseNamed } from "../04_SpreadsheetNamed/ClassBases/SheetBaseNamed";
import type { ColumnNamed } from "../04_SpreadsheetNamed/ColumnNamed";
import type { SheetNamed } from "../04_SpreadsheetNamed/SheetNamed";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import type { ConfigSyncState, OperatorProps } from "./SpreadsheetBaseOperator";

export interface SheetOperatorProps<
  TN extends TableNameSimple,
> extends OperatorProps {
  sheetName: TN;
}

export class GenericSheetOperator<
  TN extends TableNameSimple,
> extends SheetBaseNamed<TN> {
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
  get sheet(): SheetNamed<TN> {
    return this.ss.sheet(this.sheetName);
  }
  get schema(): SheetSchema<TN> {
    return SheetSchema.fromSheetName(this.sheetName);
  }
  column<CN extends ColumnName<TN>>(columnName: CN): ColumnNamed<TN, CN> {
    return this.sheet.column(columnName);
  }
}
