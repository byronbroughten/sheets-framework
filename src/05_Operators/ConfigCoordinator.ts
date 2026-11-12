import { Val } from "@byronbroughten/utils/val";

import type { IdNameMapGeneric } from "../01_SpreadsheetSchema/configReaders/idNameMapTypes";
import {
  SpreadsheetBaseNamed,
  type SpreadsheetNamedProps,
} from "../04_SpreadsheetNamed/ClassBases/SpreadsheetBaseNamed";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import { ColumnConfigOperator } from "./ColumnConfigOperator";
import { columnIdsByName, idNameMapFileSource } from "./configFileSource";
import { ConfigSheetFloor } from "./ConfigSheetFloor";
import { SpreadsheetBaseOperator } from "./SpreadsheetBaseOperator";
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
  generateConfigFiles(makeConfigsImport: string): ConfigRegeneration {
    this._syncColumnIds();
    const untypedColumnsSummary =
      this.columnConfigOperator.untypedColumnsSummary();
    this.ss.batchUpdateGSheets();
    this.valueConfigOperator.fetchAfterColumnConfigSynced();
    return {
      tableConfigs: this.tableConfigOperator.toFileSource(makeConfigsImport),
      tabConfigs: this.tabConfigOperator.toFileSource(makeConfigsImport),
      retiredIdPrefixes:
        this.tableConfigOperator.retiredIdPrefixesFileSource(makeConfigsImport),
      columnConfigs: this.columnConfigOperator.toFileSource(makeConfigsImport),
      valueConfigs: this.valueConfigOperator.toFileSource(makeConfigsImport),
      idNameMap: idNameMapFileSource(this._idNameMap()),
      untypedColumnsSummary,
    };
  }
  // The ID-prefix guard runs before the column-ID write it guards.
  private _syncColumnIds(): void {
    this.ss.fetchAllSheetProperties();
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
}
