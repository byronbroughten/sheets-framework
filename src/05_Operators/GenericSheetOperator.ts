import type { ColumnName } from "../01_SpreadsheetSchema/columnConfigsTypes";
import type { SheetNameSimple } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import { SheetSchema } from "../01_SpreadsheetSchema/SheetSchema";
import { SheetBaseNamed } from "../04_SpreadsheetNamed/ClassBases/SheetBaseNamed";
import type { ColumnNamed } from "../04_SpreadsheetNamed/ColumnNamed";
import type { SheetNamed } from "../04_SpreadsheetNamed/SheetNamed";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import type { ConfigSyncState, OperatorProps } from "./SpreadsheetBaseOperator";

export interface SheetOperatorProps<
  SN extends SheetNameSimple,
> extends OperatorProps {
  sheetName: SN;
}

export class GenericSheetOperator<
  SN extends SheetNameSimple,
> extends SheetBaseNamed<SN> {
  protected configSyncState: ConfigSyncState;
  constructor({ configSyncState, ...rest }: SheetOperatorProps<SN>) {
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
  get sheet(): SheetNamed<SN> {
    return this.ss.sheet(this.sheetName);
  }
  get schema(): SheetSchema<SN> {
    return SheetSchema.fromSheetName(this.sheetName);
  }
  column<CN extends ColumnName<SN>>(columnName: CN): ColumnNamed<SN, CN> {
    return this.sheet.column(columnName);
  }
}
