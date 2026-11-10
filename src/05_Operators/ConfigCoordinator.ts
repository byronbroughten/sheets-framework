import { Val } from "@byronbroughten/utils/val";

import type { SheetPropertiesFetchOptions } from "../00_Source/RawSource/RawSource";
import { columnConfigsByName } from "../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import { assertFloorMatchesSeed } from "../01_SpreadsheetSchema/configReaders/floorSeedCheck";
import type { IdNameMapGeneric } from "../01_SpreadsheetSchema/configReaders/idNameMapTypes";
import { tableConfigsByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import {
  SpreadsheetBaseNamed,
  type SpreadsheetNamedProps,
} from "../04_SpreadsheetNamed/ClassBases/SpreadsheetBaseNamed";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import { ColumnConfigOperator } from "./ColumnConfigOperator";
import { columnIdsByName, idNameMapFileSource } from "./configFileSource";
import { ConfigSheetFloor } from "./ConfigSheetFloor";
import { assertFloorIdentityUnchanged } from "./floorIdentityGuard";
import { SpreadsheetBaseOperator } from "./SpreadsheetBaseOperator";
import { SpreadsheetConfigOperator } from "./SpreadsheetConfigOperator";
import { TabConfigOperator } from "./TabConfigOperator";
import { TableConfigOperator } from "./TableConfigOperator";
import { ValueConfigOperator } from "./ValueConfigOperator";

export interface ConfigRegeneration {
  tableConfigs: string;
  tabConfigs: string;
  retiredIdPrefixes: string;
  columnConfigs: string;
  valueConfigs: string;
  idNameMap: string;
  untypedColumnsSummary: string | undefined;
  floorReport: string;
  idPrefixReport: string | undefined;
  declaredCellReport: string | undefined;
}

/**
 * Coordinates Spreadsheet/Table/Column Config: the config-sheet floor
 * first (one extra flush), sync the live config sheets, one more flush, then
 * emit every generated config file and the ID-name map, or none. Config maintenance is this Operator
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
    props.spreadsheetStateIdentified.isRegeneratingConfigs = true;
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
  get tabConfigOperator(): TabConfigOperator {
    return new TabConfigOperator(this.operatorProps);
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
  generateConfigFiles(): ConfigRegeneration {
    const floorReport = this._ensureFloorAndFlush();
    const untypedColumnsSummary = this._syncConfigSheetRows({
      includeConditionalFormats: true,
    });
    this.ss.batchUpdateGSheets();
    this.valueConfigOperator.fetchAfterColumnConfigSynced();
    this._assertFloorIdentityUnchanged();
    this._assertFloorMatchesSeed();
    return {
      tableConfigs: this.tableConfigOperator.toFileSource(),
      tabConfigs: this.tabConfigOperator.toFileSource(),
      retiredIdPrefixes: this.tableConfigOperator.retiredIdPrefixesFileSource(),
      columnConfigs: this.columnConfigOperator.toFileSource(),
      valueConfigs: this.valueConfigOperator.toFileSource(),
      idNameMap: idNameMapFileSource(this._idNameMap()),
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
  private _idNameMap(): IdNameMapGeneric {
    return { tables: this._idNameMapTables(), tabs: this._idNameMapTabs() };
  }
  private _idNameMapTables(): IdNameMapGeneric["tables"] {
    const { raw } = this.ss;
    const columnConfigs = this.columnConfigOperator.newColumnConfigs();
    const tableKeysByTableId = this.tableConfigOperator.tableKeysByTableId();
    return raw.activeTableIds.reduce<IdNameMapGeneric["tables"]>(
      (tables, tableId) => {
        const { name, sheetGid } = raw.table(tableId);
        const tableKey = tableKeysByTableId.get(tableId);
        if (tableKey === undefined) {
          tables[tableId] = {
            name: this.ss.schema.titleToName(name),
            sheetGid,
          };
          return tables;
        }
        tables[tableId] = {
          name: tableKey,
          sheetGid,
          columns: columnIdsByName(
            Val.assert(
              columnConfigs[tableKey],
              `column configs of ${tableKey}`,
            ),
          ),
        };
        return tables;
      },
      {},
    );
  }
  private _idNameMapTabs(): IdNameMapGeneric["tabs"] {
    const { raw } = this.ss;
    return raw.activeSheetGids.reduce<IdNameMapGeneric["tabs"]>(
      (tabs, sheetGid) => {
        const sheet = raw.sheet(sheetGid);
        tabs[String(sheetGid)] = {
          name: this.ss.schema.titleToName(sheet.title),
          title: sheet.title,
          tableIds: sheet.tableIds,
        };
        return tabs;
      },
      {},
    );
  }
  private _syncConfigSheetRows(
    propertiesFetch?: SheetPropertiesFetchOptions,
  ): string | undefined {
    this.ss.fetchAllSheetProperties(propertiesFetch);
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
