import type { IdNameMapGeneric } from "../01_SpreadsheetSchema/configReaders/idNameMapTypes";
import {
  SpreadsheetBaseNamed,
  type SpreadsheetNamedProps,
} from "../04_SpreadsheetNamed/ClassBases/SpreadsheetBaseNamed";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import { ColumnConfigsOperator } from "./ColumnConfigsOperator";
import { ConditionalFormatConfigOperator } from "./ConditionalFormatConfigOperator";
import { columnIdsByName, idNameMapFileSource } from "./configFileSource";
import { ConfigSheetFloor } from "./ConfigSheetFloor";
import { SpreadsheetBaseOperator } from "./SpreadsheetBaseOperator";
import { TabConfigOperator } from "./TabConfigOperator";
import { TableConfigsOperator } from "./TableConfigsOperator";
import { ValueConfigsOperator } from "./ValueConfigsOperator";

export interface ConfigRegeneration {
  tableConfigs: string;
  tabConfigs: string;
  conditionalFormatConfigs: string;
  retiredIdPrefixes: string;
  columnConfigs: string;
  valueConfigs: string;
  idNameMap: string;
  untypedColumnsSummary: string | undefined;
}

/**
 * Coordinates Table/Column/Value Config: reads the live Tables and the
 * previous configs, writes any missing column IDs in one flush, then emits
 * every generated config file and the ID-name map, or none. The regen leaves
 * the config tabs' rows alone; ensureConfigSheetFloor still writes them for
 * the floor chore. Config maintenance is this Operator family,
 * not Raw or Named. npm run gen:configs is the only regeneration path.
 * docs/generated-data.md
 */
export class SyncCoordinator extends SpreadsheetBaseOperator {
  constructor(props: SpreadsheetNamedProps) {
    super({
      ...props,
      configSyncState: SpreadsheetBaseOperator.initConfigSyncState(),
    });
  }
  // Generation's own state, so the sync's refusals judge the live Tables in place of the configs it replaces.
  static init(): SyncCoordinator {
    const props = SpreadsheetBaseNamed.initSpreadsheetNamedProps();
    props.spreadsheetStateIdentified.isRegeneratingConfigs = true;
    return new SyncCoordinator(props);
  }
  get ss(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  get columnConfigOperator(): ColumnConfigsOperator {
    return new ColumnConfigsOperator(this.operatorProps);
  }
  get tableConfigOperator(): TableConfigsOperator {
    return new TableConfigsOperator(this.operatorProps);
  }
  get tabConfigOperator(): TabConfigOperator {
    return new TabConfigOperator(this.operatorProps);
  }
  get conditionalFormatConfigOperator(): ConditionalFormatConfigOperator {
    return new ConditionalFormatConfigOperator(this.operatorProps);
  }
  get valueConfigOperator(): ValueConfigsOperator {
    return new ValueConfigsOperator(this.operatorProps);
  }
  get configSheetFloor(): ConfigSheetFloor {
    return new ConfigSheetFloor(this.spreadsheetNamedProps);
  }
  ensureConfigSheetFloor(): string {
    return this.configSheetFloor.ensure();
  }
  generateConfigFiles(): ConfigRegeneration {
    this._syncColumnIds();
    const untypedColumnsSummary =
      this.columnConfigOperator.untypedColumnsSummary();
    this.ss.batchUpdateGSheets();
    this.valueConfigOperator.fetchAfterColumnConfigSynced();
    return {
      tableConfigs: this.tableConfigOperator.toFileSource(),
      tabConfigs: this.tabConfigOperator.toFileSource(),
      conditionalFormatConfigs:
        this.conditionalFormatConfigOperator.toFileSource(),
      retiredIdPrefixes: this.tableConfigOperator.retiredIdPrefixesFileSource(),
      columnConfigs: this.columnConfigOperator.toFileSource(),
      valueConfigs: this.valueConfigOperator.toFileSource(),
      idNameMap: idNameMapFileSource(this._idNameMap()),
      untypedColumnsSummary,
    };
  }
  // The ID-prefix guard runs before the column-ID write it guards.
  private _syncColumnIds(): void {
    this.ss.fetchAllSheetProperties({ includeConditionalFormats: true });
    this.tableConfigOperator.validateHeadRowsClear();
    this.tableConfigOperator.validateHeadersInZone();
    this.columnConfigOperator.fetchHeadRows();
    this.tableConfigOperator.validateIdPrefixesUnchanged();
    this.columnConfigOperator.syncToSpreadsheet();
  }
  private _idNameMap(): IdNameMapGeneric {
    return { tables: this._idNameMapTables(), tabs: this._idNameMapTabs() };
  }
  private _idNameMapTables(): IdNameMapGeneric["tables"] {
    const { raw } = this.ss;
    const apiTableColumnsByTableId =
      this.columnConfigOperator.apiTableColumnsByTableId();
    return raw.activeTableIds.reduce<IdNameMapGeneric["tables"]>(
      (tables, tableId) => {
        const { name, sheetGid } = raw.table(tableId);
        const apiTableColumns = apiTableColumnsByTableId.get(tableId);
        if (apiTableColumns === undefined) {
          tables[tableId] = {
            name: this.ss.schema.titleToName(name),
            sheetGid,
          };
          return tables;
        }
        tables[tableId] = {
          name: apiTableColumns.tableKey,
          sheetGid,
          columns: columnIdsByName(apiTableColumns.columnConfigs),
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
}
