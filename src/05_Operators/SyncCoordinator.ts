import { columnConfigsByName } from "../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import { installedConfigs } from "../01_SpreadsheetSchema/configReaders/configRegister";
import type { IdNameMapGeneric } from "../01_SpreadsheetSchema/configReaders/idNameMapTypes";
import type { PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState";
import { remainingAlterations } from "../01_SpreadsheetSchema/configReaders/remainingAlterations";
import { SyncEntryRefuser } from "../01_SpreadsheetSchema/configReaders/SyncEntryRefuser";
import {
  isSheetTrait,
  type PlanProps,
  SyncPlanner,
  type TraitReadings,
} from "../01_SpreadsheetSchema/configReaders/SyncPlanner";
import type {
  SyncRunFrame,
  SyncRunResult,
  SyncSubject,
  SyncTypedTraitChange,
} from "../01_SpreadsheetSchema/configReaders/syncRunResult";
import { SyncRunResultBuilder } from "../01_SpreadsheetSchema/configReaders/SyncRunResultBuilder";
import { typedTraitChanges } from "../01_SpreadsheetSchema/configReaders/typedTraitChanges";
import {
  SpreadsheetBaseNamed,
  type SpreadsheetNamedProps,
} from "../04_SpreadsheetNamed/ClassBases/SpreadsheetBaseNamed";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import { ColumnConfigsOperator } from "./ColumnConfigsOperator";
import { ConditionalFormatConfigOperator } from "./ConditionalFormatConfigOperator";
import {
  alterationsFileSource,
  columnIdsByName,
  idNameMapFileSource,
} from "./configFileSource";
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
  // Absent unless an alteration landed, so a run that sent nothing leaves alterations.ts untouched.
  alterations?: string;
}

// Pending alterations alone until claims arrive.
export interface SyncIntendedState {
  alterations: PendingAlterations;
}

export interface SyncRunProps {
  // A bare run plans but sends nothing; --send sends.
  isSending: boolean;
  intendedState: SyncIntendedState;
  previousIdNameMap: IdNameMapGeneric;
  frame: SyncRunFrame;
}

interface Regeneration {
  files: ConfigRegeneration;
  changes: SyncTypedTraitChange[];
  notes: string[];
}

// No files on a refused run: the run writes every file or none.
export interface SyncRun {
  result: SyncRunResult;
  files?: ConfigRegeneration;
}

/**
 * Runs sync: refuses what intended state alone can't plan, reads the live
 * Tables, plans the three-way comparison against the previous configs, sends
 * the plan and any missing column IDs in one flush on an unstopped --send,
 * refetches, then emits every generated config file, the ID-name map and the
 * remaining alterations, or none. A bare run sends nothing but still
 * refreshes the configs from its read. The run leaves the config tabs' rows
 * alone; ensureConfigSheetFloor still writes them for the floor chore.
 * Config maintenance is this Operator family, not Raw or Named.
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
  run(props: SyncRunProps): SyncRun {
    const { alterations } = props.intendedState;
    const planner = SyncPlanner.init({
      alterations,
      tabsByTable: tabsByTable(props.previousIdNameMap),
      forcedTraits: [],
    });
    const refusals = [
      ...SyncEntryRefuser.init({
        alterations,
        idNameMap: props.previousIdNameMap,
      }).refusals,
      ...planner.refusals,
    ];
    const resultBuilder = SyncRunResultBuilder.init(props);
    if (refusals.length > 0) {
      return { result: resultBuilder.refused(refusals) };
    }
    this._readLive();
    // The ID-prefix guard runs before the column-ID write it guards.
    this.tableConfigOperator.validateIdPrefixesUnchanged();
    const readings = this._traitReadings();
    const planned = planner.plan(readings);
    // The gate's holds go here, between the plan and the apply.
    const isApplying = props.isSending && !planned.isStopped;
    // An unstopped dry run records the live readings: configTraits carry sends it won't make.
    this.columnConfigOperator.recordEmptyValueAllowed(
      props.isSending || planned.isStopped
        ? planned.configTraits
        : readings.live,
    );
    // Column IDs only on an unstopped send: one minted by any other run would reach the configs but never the sheet.
    if (isApplying) {
      this.columnConfigOperator
        .addMissingColumnIds()
        .updateHeaders(planned.sentTraits);
      this.tableConfigOperator.updateTableNames(planned.sentTraits);
      this.tabConfigOperator.updateTabTraits(planned.sentTraits);
    }
    const notes = this._recordUntypedColumns();
    if (isApplying) this.ss.batchUpdateGSheets();
    const sent = isApplying ? planned.plan.map(({ subject }) => subject) : [];
    // A fresh read, since Raw's state has no reset; only a sent sheet trait changes what it would read.
    const regeneration =
      isApplying && planned.sentTraits.some(isSheetTrait)
        ? SyncCoordinator.init()._refetchedRegeneration(
            planned.configTraits,
            sent,
          )
        : this._regeneration(notes, sent);
    const { files, changes } = regeneration;
    // The row-ID fill goes here: its own send, planned from the fresh configs.
    if (isApplying && planned.plan.length > 0) {
      files.alterations = alterationsFileSource(
        remainingAlterations(
          alterations,
          planned.plan.map(({ subject }) => subject),
        ),
      );
    }
    return {
      result: resultBuilder.ran({
        planned,
        notes: regeneration.notes,
        typedTraitChanges: changes,
      }),
      files,
    };
  }
  // The untyped-column notes come from the refetch, since a rename changes the headers and Table names they cite.
  private _refetchedRegeneration(
    configTraits: TraitReadings,
    sent: SyncSubject[],
  ): Regeneration {
    this._readLive();
    this.columnConfigOperator.recordEmptyValueAllowed(configTraits);
    return this._regeneration(this._recordUntypedColumns(), sent);
  }
  private _regeneration(notes: string[], sent: SyncSubject[]): Regeneration {
    this.valueConfigOperator.fetchAfterColumnConfigSynced();
    const changes = typedTraitChanges({
      previous: {
        columnConfigs: columnConfigsByName(),
        valueConfigs: installedConfigs().valueConfigs,
      },
      regenerated: {
        columnConfigs: this.columnConfigOperator.newColumnConfigs(),
        valueConfigs: this.valueConfigOperator.newValueConfigs(),
      },
      sent,
    });
    return { files: this._configFiles(), changes, notes };
  }
  private _traitReadings(): PlanProps {
    const { previous, live } = this.columnConfigOperator.traitReadings();
    const tabs = this.tabConfigOperator.traitReadings();
    return {
      previous: { ...previous, tabs: tabs.previous },
      live: { ...live, tabs: tabs.live },
    };
  }
  private _readLive(): void {
    this.ss.fetchAllSheetProperties({ includeConditionalFormats: true });
    this.tableConfigOperator.validateHeadRowsClear();
    this.tableConfigOperator.validateHeadersInZone();
    this.columnConfigOperator.fetchHeadRows();
  }
  private _recordUntypedColumns(): string[] {
    const untypedColumnsSummary = this.columnConfigOperator
      .recordUntypedColumns()
      .untypedColumnsSummary();
    if (untypedColumnsSummary === undefined) return [];
    return [untypedColumnsSummary];
  }
  private _configFiles(): ConfigRegeneration {
    return {
      tableConfigs: this.tableConfigOperator.toFileSource(),
      tabConfigs: this.tabConfigOperator.toFileSource(),
      conditionalFormatConfigs:
        this.conditionalFormatConfigOperator.toFileSource(),
      retiredIdPrefixes: this.tableConfigOperator.retiredIdPrefixesFileSource(),
      columnConfigs: this.columnConfigOperator.toFileSource(),
      valueConfigs: this.valueConfigOperator.toFileSource(),
      idNameMap: idNameMapFileSource(this._idNameMap()),
    };
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

// Each Table's tab by name, from the previous map, since nothing is read before the refusals.
function tabsByTable({
  tables,
  tabs,
}: IdNameMapGeneric): Record<string, string | undefined> {
  return Object.values(tables).reduce<Record<string, string | undefined>>(
    (byTable, { name, sheetGid }) => {
      byTable[name] = tabs[String(sheetGid)]?.name;
      return byTable;
    },
    {},
  );
}
