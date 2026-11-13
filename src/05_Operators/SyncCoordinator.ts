import { HeldFlushRefusal } from "../00_Source/RawSource/HeldFlushRefusal";
import {
  type ApiAccessChanges,
  apiAccessChanges,
  withApiAccessReadings,
} from "../01_SpreadsheetSchema/configReaders/apiAccessChanges";
import { columnConfigsByName } from "../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import { installedConfigs } from "../01_SpreadsheetSchema/configReaders/configRegister";
import {
  didntLandOf,
  hasLanded,
} from "../01_SpreadsheetSchema/configReaders/didntLand";
import {
  type EnteredDataReads,
  enteredDataReads,
} from "../01_SpreadsheetSchema/configReaders/enteredDataReads";
import type { IdNameMapGeneric } from "../01_SpreadsheetSchema/configReaders/idNameMapTypes";
import type { PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState";
import { remainingAlterations } from "../01_SpreadsheetSchema/configReaders/remainingAlterations";
import { SyncEntryRefuser } from "../01_SpreadsheetSchema/configReaders/SyncEntryRefuser";
import {
  isSheetTrait,
  type PlanProps,
  SyncPlanner,
  type SyncPlannerResult,
  type TraitReadings,
} from "../01_SpreadsheetSchema/configReaders/SyncPlanner";
import type {
  ForcedTrait,
  SyncPlanItem,
  SyncRunFrame,
  SyncRunResult,
  SyncSendLine,
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
import { RowIdsOperator } from "./RowIdsOperator";
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
  // Each --force flag of this run alone; never read from a file.
  forcedTraits: ForcedTrait[];
  // Set by the Node half, whose host withholds a dry run's flushes, so an unstopped dry run flushes for the gate to judge.
  isHostWithholding?: boolean;
}

interface Regeneration {
  files: ConfigRegeneration;
  changes: SyncTypedTraitChange[];
  notes: string[];
  didntLand: SyncSendLine[];
}

// What a plan's flush writes, and the cells read before it so the gate can judge them.
interface AppliedPlan {
  accessChanges: ApiAccessChanges;
  sentTraits: SyncPlannerResult["sentTraits"];
  reads: EnteredDataReads;
}

interface RefetchProps {
  configTraits: TraitReadings;
  plan: SyncPlannerResult;
  sent: SyncSubject[];
  accessChanges: ApiAccessChanges;
}

// No files on a refused run: the run writes every file or none.
export interface SyncRun {
  result: SyncRunResult;
  files?: ConfigRegeneration;
}

/**
 * Runs sync: refuses what intended state alone can't plan, reads the live
 * Tables, plans the three-way comparison against the previous configs, sends
 * the plan, any missing column IDs and Fill Row IDs checkboxes in one flush
 * on an unstopped --send, refetches, then emits every generated config file,
 * the ID-name map and the remaining alterations, or none, and fills blank
 * row IDs from the fresh configs as its own send. A bare run sends nothing but still
 * refreshes the configs from its read; on a host that withholds, it then
 * flushes the plan for the host's entered-data gate to judge. Before either
 * flush it reads the cells a deletion or a formula would overwrite, so the
 * gate judges them rather than holding them unread. A send the
 * gate refuses writes no files. The run leaves the config tabs' rows
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
  get rowIdsOperator(): RowIdsOperator {
    return new RowIdsOperator(this.operatorProps);
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
      forcedTraits: props.forcedTraits,
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
    const accessChanges = apiAccessChanges({
      alterations,
      idNameMap: props.previousIdNameMap,
    });
    this._readLive();
    // The ID-prefix guard runs before the column-ID write it guards.
    this.tableConfigOperator.validateIdPrefixesUnchanged();
    this.rowIdsOperator.validateIdPrefixes();
    const checkboxPlan = this.rowIdsOperator.checkboxPlan();
    const notices = this.rowIdsOperator.idFirstNotices();
    const readings = this._traitReadings();
    const planned = planner.plan({
      ...withApiAccessReadings(readings, props.previousIdNameMap),
      headRowStops: this.tableConfigOperator.grantHeadRowStops(
        accessChanges.grants,
      ),
    });
    const isApplying = props.isSending && !planned.isStopped;
    // An unstopped dry run records the live readings: configTraits carry sends it won't make.
    this.columnConfigOperator.recordEmptyValueAllowed(
      props.isSending || planned.isStopped
        ? planned.configTraits
        : readings.live,
    );
    const applied: AppliedPlan = {
      accessChanges,
      sentTraits: planned.sentTraits,
      reads: enteredDataReads({
        alterations,
        sentTraits: planned.sentTraits,
        idNameMap: props.previousIdNameMap,
      }),
    };
    // Grants join here only on an unstopped send; a dry run applies them after its files are built, so its configs never record a Table with no column IDs.
    // Column IDs only on an unstopped send: one minted by any other run would reach the configs but never the sheet.
    if (isApplying) this._applyPlan(applied);
    const notes = this._recordUntypedColumns();
    if (isApplying) {
      const refusal = this._flushUnlessHeld();
      // No files: column IDs already recorded would reach the configs but not the sheet.
      if (refusal !== undefined) {
        return {
          result: resultBuilder.heldStopped({
            planned,
            refusal,
            rowIds: { plan: checkboxPlan, notices },
          }),
        };
      }
    }
    const sent = isApplying ? planned.plan.map(({ subject }) => subject) : [];
    let refetched: SyncCoordinator | undefined;
    let regeneration: Regeneration;
    // A fresh read, since Raw's state has no reset; only a sent sheet trait changes what it would read.
    if (isApplying && planned.sentTraits.some(isSheetTrait)) {
      refetched = SyncCoordinator.init();
      regeneration = refetched._refetchedRegeneration({
        configTraits: planned.configTraits,
        plan: planned,
        sent,
        accessChanges,
      });
    } else {
      regeneration = this._regeneration(notes, sent);
    }
    // The fill reads the configs the files came from.
    const regenerator = refetched ?? this;
    const { files, changes, didntLand } = regeneration;
    // After the files are strings, so the host's gate judges what a send would write and the configs never see it.
    if (props.isHostWithholding && !props.isSending && !planned.isStopped) {
      this._applyPlan(applied);
      this.ss.batchUpdateGSheets();
    }
    const landed = sent.filter(hasLanded(didntLand));
    if (landed.length > 0) {
      files.alterations = alterationsFileSource(
        remainingAlterations(alterations, landed),
      );
    }
    let fillPlan: SyncPlanItem[];
    if (isApplying) {
      fillPlan = regenerator._fillBlankIds();
    } else {
      fillPlan = regenerator.rowIdsOperator.fillPlan();
    }
    return {
      result: resultBuilder.ran({
        planned,
        notes: regeneration.notes,
        typedTraitChanges: changes,
        didntLand,
        rowIds: { plan: [...checkboxPlan, ...fillPlan], notices },
      }),
      files,
    };
  }
  // Its own send after the regen, planned from the fresh configs; it writes only blank cells, so the entered-data gate holds nothing.
  private _fillBlankIds(): SyncPlanItem[] {
    const filled = this.rowIdsOperator.fillBlankIds();
    this.ss.batchUpdateGSheets();
    return filled;
  }
  // The untyped-column notes come from the refetch, since a rename changes the headers and Table names they cite.
  private _refetchedRegeneration({
    configTraits,
    plan,
    sent,
    accessChanges,
  }: RefetchProps): Regeneration {
    // Recorded before the read, so its head-row fetch covers granted Tables and skips revoked ones.
    this.tableConfigOperator.recordApiAccessChanges(accessChanges);
    this._readLive();
    this.columnConfigOperator.recordEmptyValueAllowed(configTraits);
    const didntLand = didntLandOf({
      sentTraits: plan.sentTraits,
      plan: plan.plan,
      live: this._traitReadings().live,
    });
    return this._regeneration(this._recordUntypedColumns(), sent, didntLand);
  }
  private _regeneration(
    notes: string[],
    sent: SyncSubject[],
    didntLand: SyncSendLine[] = [],
  ): Regeneration {
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
    return { files: this._configFiles(), changes, notes, didntLand };
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
  private _applyPlan({ accessChanges, sentTraits, reads }: AppliedPlan): void {
    this._readEnteredData(reads);
    // Before the access changes, so it sends what the pre-send read planned.
    this.rowIdsOperator.addCheckboxes();
    this._applyApiAccessChanges(accessChanges);
    this.columnConfigOperator
      .addMissingColumnIds()
      .updateHeaders(sentTraits)
      .updateFormulas(sentTraits);
    this.tableConfigOperator.updateTableNames(sentTraits);
    this.tabConfigOperator.updateTabTraits(sentTraits);
  }
  // The cells the host's entered-data gate judges; a Table or tab gone from the sheet has none to read.
  private _readEnteredData({
    tabGids,
    tableIds,
    columns,
  }: EnteredDataReads): void {
    const { raw, identified } = this.ss;
    for (const sheetGid of tabGids) {
      if (raw.gidIsActive(sheetGid)) raw.fetchSheetUsedGrid(sheetGid);
    }
    for (const tableId of tableIds) {
      if (!raw.tableIdIsActive(tableId)) continue;
      const table = raw.table(tableId);
      for (const colIndex of table.fullTableColIndexes) {
        table.column(colIndex).gatherFetchFull();
      }
    }
    for (const { tableId, columnId } of columns) {
      if (!raw.tableIdIsActive(tableId)) continue;
      const table = identified.table(tableId);
      if (!table.columnResolver.hasColumnId(columnId)) continue;
      table.column(columnId).raw.gatherFetchFull();
    }
    identified.fetchAllGathered();
  }
  // The Node host's gate refuses a held flush by throwing; any other error goes on up.
  private _flushUnlessHeld(): HeldFlushRefusal | undefined {
    try {
      this.ss.batchUpdateGSheets();
      return undefined;
    } catch (error) {
      if (error instanceof HeldFlushRefusal) return error;
      throw error;
    }
  }
  private _applyApiAccessChanges(accessChanges: ApiAccessChanges): void {
    this.tableConfigOperator.recordApiAccessChanges(accessChanges);
    if (accessChanges.grants.length === 0) return;
    this.tableConfigOperator.validateHeadRowsClear();
    this.columnConfigOperator.fetchHeadRowsOf(
      accessChanges.grants.map(({ tableId }) => tableId),
    );
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
