import {
  SpreadsheetBaseNamed,
  type SpreadsheetNamedProps,
} from "../04_SpreadsheetNamed/ClassBases/SpreadsheetBaseNamed";

export type UntypedHeadersBySheetTitle = Map<string, string[]>;

// Lives on operator props, not each collaborator, so getter rebuilds share it.
export interface ConfigSyncState {
  sheetConfigSync: {
    prepFetchIsComplete: boolean;
    syncedToSpreadsheet: boolean;
    declaredCellReportLines: string[];
  };
  columnConfigSync: {
    syncedToSpreadsheet: boolean;
    untypedHeadersBySheetTitle: UntypedHeadersBySheetTitle;
    declaredCellReportLines: string[];
  };
  valueConfigSync: { activeHeaders: Set<string> };
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
      sheetConfigSync: {
        prepFetchIsComplete: false,
        syncedToSpreadsheet: false,
        declaredCellReportLines: [],
      },
      columnConfigSync: {
        syncedToSpreadsheet: false,
        untypedHeadersBySheetTitle: new Map(),
        declaredCellReportLines: [],
      },
      valueConfigSync: { activeHeaders: new Set() },
    };
  }
  static initOperatorProps(): OperatorProps {
    return {
      ...SpreadsheetBaseNamed.initSpreadsheetNamedProps(),
      configSyncState: SpreadsheetBaseOperator.initConfigSyncState(),
    };
  }
}
