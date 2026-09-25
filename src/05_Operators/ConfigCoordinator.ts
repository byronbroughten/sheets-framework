import { columnConfigsByName } from "../01_SpreadsheetSchema/columnConfigsTypes";
import { assertFloorMatchesSeed } from "../01_SpreadsheetSchema/floorSeedCheck";
import { sheetConfigsByName } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import {
  clearSpreadsheetConfigOverlay,
  overlaySpreadsheetConfig,
} from "../01_SpreadsheetSchema/spreadsheetConfigTypes";
import {
  SpreadsheetBaseNamed,
  type SpreadsheetNamedProps,
} from "../04_SpreadsheetNamed/ClassBases/SpreadsheetBaseNamed";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import { ColumnConfigOperator } from "./ColumnConfigOperator";
import { ConfigSheetFloor } from "./ConfigSheetFloor";
import { assertFloorIdentityUnchanged } from "./floorIdentityGuard";
import { SheetConfigOperator } from "./SheetConfigOperator";
import { SpreadsheetBaseOperator } from "./SpreadsheetBaseOperator";
import { SpreadsheetConfigOperator } from "./SpreadsheetConfigOperator";
import { ValueConfigOperator } from "./ValueConfigOperator";

export interface ConfigRegeneration {
  spreadsheetConfig: string;
  sheetConfigs: string;
  columnConfigs: string;
  valueConfigs: string;
  untypedColumnsSummary: string | undefined;
  floorReport: string;
  idPrefixReport: string | undefined;
  declaredCellReport: string | undefined;
}

/**
 * Coordinates Spreadsheet/Sheet/Column/Value Config: the config-sheet floor
 * first (one extra flush), overlay live layout, sync the live config sheets,
 * one more flush, then emit all four generated files or none. Config
 * maintenance is this Operator family, not Raw or Named. npm run gen:configs
 * is the only regeneration path.
 * docs/generated-data.md
 */
export class ConfigCoordinator extends SpreadsheetBaseOperator {
  constructor(props: SpreadsheetNamedProps) {
    super({
      ...props,
      configSyncState: SpreadsheetBaseOperator.initConfigSyncState(),
    });
  }
  static init(): ConfigCoordinator {
    return new ConfigCoordinator(
      SpreadsheetBaseNamed.initSpreadsheetNamedProps(),
    );
  }
  get ss(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  get spreadsheetConfigOperator(): SpreadsheetConfigOperator {
    return new SpreadsheetConfigOperator(this.operatorProps);
  }
  get columnConfigOperator(): ColumnConfigOperator {
    return new ColumnConfigOperator(this.operatorProps);
  }
  get sheetConfigOperator(): SheetConfigOperator {
    return new SheetConfigOperator(this.operatorProps);
  }
  get valueConfigOperator(): ValueConfigOperator {
    return new ValueConfigOperator(this.operatorProps);
  }
  get configSheetFloor(): ConfigSheetFloor {
    return new ConfigSheetFloor(this.spreadsheetNamedProps);
  }
  ensureConfigSheetFloor(): string {
    return this.configSheetFloor.ensure();
  }
  // Returns the run status an endpoint should report, if there's one to make.
  syncConfigSheetRows(): string | undefined {
    return this._withFloorThenLiveConfig((floorReport) =>
      this._combinedSyncReport(floorReport),
    );
  }
  syncAndFlushConfigSheets(): string | undefined {
    return this._withFloorThenLiveConfig((floorReport) => {
      const summary = this._combinedSyncReport(floorReport);
      this.ss.batchUpdateGSheets();
      return summary;
    });
  }
  generateConfigFiles(makeConfigsImport: string): ConfigRegeneration {
    return this._withFloorThenLiveConfig((floorReport) => {
      const untypedColumnsSummary = this._syncConfigSheetRows();
      this.ss.batchUpdateGSheets();
      this.valueConfigOperator.fetchAfterColumnConfigSynced();
      this._assertFloorIdentityUnchanged();
      this._assertFloorMatchesSeed();
      return {
        spreadsheetConfig:
          this.spreadsheetConfigOperator.toFileSource(makeConfigsImport),
        sheetConfigs: this.sheetConfigOperator.toFileSource(makeConfigsImport),
        columnConfigs:
          this.columnConfigOperator.toFileSource(makeConfigsImport),
        valueConfigs: this.valueConfigOperator.toFileSource(makeConfigsImport),
        untypedColumnsSummary,
        floorReport,
        idPrefixReport: this.sheetConfigOperator.idPrefixChangeReport(),
        declaredCellReport: this._declaredCellReport(),
      };
    });
  }
  private _withFloorThenLiveConfig<RT>(body: (floorReport: string) => RT): RT {
    const floorReport = this.ensureConfigSheetFloor();
    this.ss.batchUpdateGSheets();
    return this._withLiveSpreadsheetConfig(() => body(floorReport));
  }
  private _withLiveSpreadsheetConfig<RT>(body: () => RT): RT {
    const liveConfig = this.spreadsheetConfigOperator.fetchLiveConfig();
    overlaySpreadsheetConfig(liveConfig);
    try {
      return body();
    } finally {
      clearSpreadsheetConfigOverlay();
    }
  }
  private _assertFloorIdentityUnchanged(): void {
    assertFloorIdentityUnchanged({
      previous: {
        sheetConfigs: sheetConfigsByName(),
        columnConfigs: columnConfigsByName(),
      },
      next: {
        sheetConfigs: this.sheetConfigOperator.newSheetConfigs(),
        columnConfigs: this.columnConfigOperator.newColumnConfigs(),
      },
    });
  }
  private _assertFloorMatchesSeed(): void {
    assertFloorMatchesSeed(
      this.sheetConfigOperator.newSheetConfigs(),
      this.columnConfigOperator.newColumnConfigs(),
    );
  }
  private _syncConfigSheetRows(): string | undefined {
    this.ss.fetchAllSheetProperties();
    this.spreadsheetConfigOperator.validateExactlyOneDataRow();
    this.sheetConfigOperator.prepFetchForSync();
    this.columnConfigOperator.prepFetchWithSheetConfig();
    this.ss.fetchAllPrepped({ skipFetchingProperties: true });
    this.sheetConfigOperator.syncToSpreadsheet();
    this.columnConfigOperator.fetchAfterSheetConfigSynced();
    this.columnConfigOperator.syncToSpreadsheet();
    return this.columnConfigOperator.untypedColumnsSummary();
  }
  private _combinedSyncReport(floorReport: string): string | undefined {
    const untypedColumnsSummary = this._syncConfigSheetRows();
    return combineConfigSyncReports(
      floorReport,
      this._declaredCellReport(),
      untypedColumnsSummary,
    );
  }
  private _declaredCellReport(): string | undefined {
    return combineConfigSyncReports(
      this.sheetConfigOperator.declaredCellReport(),
      this.columnConfigOperator.declaredCellReport(),
    );
  }
}

function combineConfigSyncReports(
  ...parts: Array<string | undefined>
): string | undefined {
  const present = parts.filter(
    (part): part is string => part !== undefined && part !== "",
  );
  if (present.length === 0) return undefined;
  return present.join(" ");
}
