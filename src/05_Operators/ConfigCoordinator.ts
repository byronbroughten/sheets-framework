import { columnConfigsByName } from "../01_SpreadsheetSchema/columnConfigsTypes";
import { assertFloorMatchesSeed } from "../01_SpreadsheetSchema/floorSeedCheck";
import { tableConfigsByName } from "../01_SpreadsheetSchema/tableConfigsTypes";
import {
  SpreadsheetBaseNamed,
  type SpreadsheetNamedProps,
} from "../04_SpreadsheetNamed/ClassBases/SpreadsheetBaseNamed";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import { ColumnConfigOperator } from "./ColumnConfigOperator";
import { ConfigSheetFloor } from "./ConfigSheetFloor";
import { assertFloorIdentityUnchanged } from "./floorIdentityGuard";
import { SpreadsheetBaseOperator } from "./SpreadsheetBaseOperator";
import { SpreadsheetConfigOperator } from "./SpreadsheetConfigOperator";
import { TableConfigOperator } from "./TableConfigOperator";
import { ValueConfigOperator } from "./ValueConfigOperator";

export interface ConfigRegeneration {
  tableConfigs: string;
  columnConfigs: string;
  valueConfigs: string;
  untypedColumnsSummary: string | undefined;
  floorReport: string;
  idPrefixReport: string | undefined;
  declaredCellReport: string | undefined;
}

/**
 * Coordinates Spreadsheet/Table/Column/Value Config: the config-sheet floor
 * first (one extra flush), sync the live config sheets, one more flush, then
 * emit all three generated files or none. Config maintenance is this Operator
 * family, not Raw or Named. npm run gen:configs is the only regeneration path.
 * docs/generated-data.md
 */
export class ConfigCoordinator extends SpreadsheetBaseOperator {
  constructor(props: SpreadsheetNamedProps) {
    super({
      ...props,
      configSyncState: SpreadsheetBaseOperator.initConfigSyncState(),
    });
  }
  // Generation's own state, so the sync's refusals judge the live Tables in place of the configs it replaces.
  static init(): ConfigCoordinator {
    const props = SpreadsheetBaseNamed.initSpreadsheetNamedProps();
    props.spreadsheetStateRaw.isRegeneratingConfigs = true;
    return new ConfigCoordinator(props);
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
  get tableConfigOperator(): TableConfigOperator {
    return new TableConfigOperator(this.operatorProps);
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
    return this._combinedSyncReport(this._ensureFloorAndFlush());
  }
  syncAndFlushConfigSheets(): string | undefined {
    const summary = this.syncConfigSheetRows();
    this.ss.batchUpdateGSheets();
    return summary;
  }
  generateConfigFiles(makeConfigsImport: string): ConfigRegeneration {
    const floorReport = this._ensureFloorAndFlush();
    const untypedColumnsSummary = this._syncConfigSheetRows();
    this.ss.batchUpdateGSheets();
    this.valueConfigOperator.fetchAfterColumnConfigSynced();
    this._assertFloorIdentityUnchanged();
    this._assertFloorMatchesSeed();
    return {
      tableConfigs: this.tableConfigOperator.toFileSource(makeConfigsImport),
      columnConfigs: this.columnConfigOperator.toFileSource(makeConfigsImport),
      valueConfigs: this.valueConfigOperator.toFileSource(makeConfigsImport),
      untypedColumnsSummary,
      floorReport,
      idPrefixReport: this.tableConfigOperator.idPrefixChangeReport(),
      declaredCellReport: this._declaredCellReport(),
    };
  }
  private _ensureFloorAndFlush(): string {
    const floorReport = this.ensureConfigSheetFloor();
    this.ss.batchUpdateGSheets();
    return floorReport;
  }
  private _assertFloorIdentityUnchanged(): void {
    assertFloorIdentityUnchanged({
      previous: {
        tableConfigs: tableConfigsByName(),
        columnConfigs: columnConfigsByName(),
      },
      next: {
        tableConfigs: this.tableConfigOperator.newTableConfigs(),
        columnConfigs: this.columnConfigOperator.newColumnConfigs(),
      },
    });
  }
  private _assertFloorMatchesSeed(): void {
    assertFloorMatchesSeed(
      this.tableConfigOperator.newTableConfigs(),
      this.columnConfigOperator.newColumnConfigs(),
    );
  }
  private _syncConfigSheetRows(): string | undefined {
    this.ss.fetchAllSheetProperties();
    this.spreadsheetConfigOperator.validateExactlyOneDataRow();
    this.tableConfigOperator.prepFetchForSync();
    this.columnConfigOperator.prepFetchWithTableConfig();
    this.ss.fetchAllPrepped({ skipFetchingProperties: true });
    this.tableConfigOperator.syncToSpreadsheet();
    this.tableConfigOperator.validateHeadRowsClear();
    this.tableConfigOperator.validateHeadersInZone();
    this.columnConfigOperator.fetchAfterTableConfigSynced();
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
      this.tableConfigOperator.declaredCellReport(),
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
