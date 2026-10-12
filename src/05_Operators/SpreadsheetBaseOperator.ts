import type { TableColumnReference } from "../02_SpreadsheetRaw/ColumnProfileRaw";
import {
  SpreadsheetBaseNamed,
  type SpreadsheetNamedProps,
} from "../04_SpreadsheetNamed/ClassBases/SpreadsheetBaseNamed";

export type UntypedHeadersByTableName = Map<string, string[]>;

export interface SourceColumn {
  tableId: string;
  header: string;
}
export interface ValueSource extends SourceColumn {
  valueName: string;
}

// Lives on operator props, not each collaborator, so getter rebuilds share it.
export interface ConfigSyncState {
  tableConfigSync: {
    prepFetchIsComplete: boolean;
    syncedToSpreadsheet: boolean;
    declaredCellReportLines: string[];
  };
  columnConfigSync: {
    syncedToSpreadsheet: boolean;
    untypedHeadersByTableName: UntypedHeadersByTableName;
    declaredCellReportLines: string[];
  };
  valueConfigSync: { valueSources: ValueSource[] };
}

export interface OperatorProps extends SpreadsheetNamedProps {
  configSyncState: ConfigSyncState;
}

export class SpreadsheetBaseOperator extends SpreadsheetBaseNamed {
  protected configSyncState: ConfigSyncState;
  constructor({ configSyncState, ...rest }: OperatorProps) {
    super(rest);
    this.configSyncState = configSyncState;
  }
  get operatorProps(): OperatorProps {
    return {
      ...this.spreadsheetNamedProps,
      configSyncState: this.configSyncState,
    };
  }
  static initConfigSyncState(): ConfigSyncState {
    return {
      tableConfigSync: {
        prepFetchIsComplete: false,
        syncedToSpreadsheet: false,
        declaredCellReportLines: [],
      },
      columnConfigSync: {
        syncedToSpreadsheet: false,
        untypedHeadersByTableName: new Map(),
        declaredCellReportLines: [],
      },
      valueConfigSync: { valueSources: [] },
    };
  }
  static initOperatorProps(): OperatorProps {
    return {
      ...SpreadsheetBaseNamed.initSpreadsheetNamedProps(),
      configSyncState: SpreadsheetBaseOperator.initConfigSyncState(),
    };
  }
}

export function sourceLabel({ tableName, header }: TableColumnReference): string {
  return `"${tableName}[${header}]"`;
}
