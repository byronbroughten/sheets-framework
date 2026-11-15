import { Val } from "@byronbroughten/utils/val";

import { HeldFlushRefusal } from "../00_Source/RawSource/HeldFlushRefusal";
import { PartialTableRefusal } from "../00_Source/RawSource/PartialTableRefusal";
import {
  type ApiAccessChanges,
  apiAccessChanges,
  withApiAccessReadings,
} from "../01_SpreadsheetSchema/configReaders/apiAccessChanges";
import { columnConfigsByName } from "../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import { installedConfigs } from "../01_SpreadsheetSchema/configReaders/configRegister";
import type {
  ConfigSnapshot,
  ConfigSnapshots,
} from "../01_SpreadsheetSchema/configReaders/configSnapshots";
import {
  didntLandOf,
  hasLanded,
} from "../01_SpreadsheetSchema/configReaders/didntLand";
import {
  type EnteredDataReads,
  enteredDataReads,
} from "../01_SpreadsheetSchema/configReaders/enteredDataReads";
import {
  heldStoppedOutcome,
  refusedStoppedOutcome,
} from "../01_SpreadsheetSchema/configReaders/flushStoppedOutcome";
import {
  type IdNameMapGeneric,
  tableIdsByName,
} from "../01_SpreadsheetSchema/configReaders/idNameMapTypes";
import type { PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState";
import { remainingAlterations } from "../01_SpreadsheetSchema/configReaders/remainingAlterations";
import { SyncEntryRefuser } from "../01_SpreadsheetSchema/configReaders/SyncEntryRefuser";
import { SyncNameRewriter } from "../01_SpreadsheetSchema/configReaders/SyncNameRewriter";
import {
  deletedTableOf,
  isSheetTrait,
  type PlanProps,
  SyncPlanner,
  type SyncPlannerResult,
  type TraitReadings,
} from "../01_SpreadsheetSchema/configReaders/SyncPlanner";
import type {
  ForcedTrait,
  SyncCreatedColumn,
  SyncFlushStoppedOutcome,
  SyncNotice,
  SyncPlanItem,
  SyncRunFrame,
  SyncRunResult,
  SyncSendLine,
  SyncSubject,
  SyncTabGoneNotice,
  SyncTableGoneNotice,
  SyncTypedTraitChange,
} from "../01_SpreadsheetSchema/configReaders/syncRunResult";
import { SyncRunResultBuilder } from "../01_SpreadsheetSchema/configReaders/SyncRunResultBuilder";
import { typedTraitChanges } from "../01_SpreadsheetSchema/configReaders/typedTraitChanges";
import {
  SpreadsheetBaseNamed,
  type SpreadsheetNamedProps,
} from "../04_SpreadsheetNamed/ClassBases/SpreadsheetBaseNamed";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import {
  ColumnConfigsOperator,
  type MintedNewColumn,
  type PlannedColumnMove,
} from "./ColumnConfigsOperator";
import {
  type CompiledRuleReplacement,
  ConditionalFormatConfigOperator,
  type IndexedRuleRemoval,
} from "./ConditionalFormatConfigOperator";
import {
  alterationsFileSource,
  columnIdsByName,
  idNameMapFileSource,
} from "./configFileSource";
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
  // The configs the files record, beside the ones the run started from.
  configs: ConfigSnapshots;
  changes: SyncTypedTraitChange[];
  notes: string[];
  didntLand: SyncSendLine[];
  createdColumns: SyncCreatedColumn[];
}

// What a plan's flush writes, and the cells read before it so the gate can judge them.
interface AppliedPlan {
  accessChanges: ApiAccessChanges;
  sentTraits: SyncPlannerResult["sentTraits"];
  newColumns: MintedNewColumn[];
  columnMoves: PlannedColumnMove[];
  ruleRemovals: IndexedRuleRemoval[];
  ruleReplacements: CompiledRuleReplacement[];
  reads: EnteredDataReads;
}

interface RefetchProps {
  configTraits: TraitReadings;
  plan: SyncPlannerResult;
  sent: SyncSubject[];
  previous: ConfigSnapshot;
  accessChanges: ApiAccessChanges;
  newColumns: MintedNewColumn[];
  ruleRemovals: IndexedRuleRemoval[];
  ruleReplacements: CompiledRuleReplacement[];
}

interface RegenerationProps {
  notes: string[];
  sent: SyncSubject[];
  previous: ConfigSnapshot;
}

// No files on a refused run: the run writes every file or none.
export interface SyncRun {
  result: SyncRunResult;
  files?: ConfigRegeneration;
}

/**
 * Runs sync: refuses what intended state alone can't plan, reads the live
 * Tables, plans the three-way comparison against the previous configs and
 * each new column, stopping on a new column whose header its Table already
 * has. On an unstopped --send it mints each new column's ID, sends the plan,
 * any missing column IDs, Fill Row IDs checkboxes and new columns at a
 * Table's end in one flush, then its column deletions and the new columns
 * placed before another in a flush of their own,
 * refetches, then emits every generated config file, the ID-name map and
 * the remaining alterations, or none, and fills blank
 * row IDs from the fresh configs as its own send. Every run that refreshes the
 * ID-name map rewrites the pending entries naming a renamed Table, column or
 * tab, a bare or stopped run included. A bare run sends nothing but still
 * refreshes the configs from its read; on a host that withholds, it then
 * flushes the plan for the host's entered-data gate to judge. Before either
 * flush it reads the cells a deletion or a formula would overwrite, so the
 * gate judges them rather than holding them unread. A send the
 * gate refuses writes no files.
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
  run(props: SyncRunProps): SyncRun {
    const { alterations } = props.intendedState;
    const tabOfTable = tabsByTable(props.previousIdNameMap);
    const planner = SyncPlanner.init({
      alterations,
      tabsByTable: tabOfTable,
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
    const notices = [
      ...goneNotices({
        tabsGone: this.tabConfigOperator.goneTabNotices(
          props.previousIdNameMap,
        ),
        tablesGone: this.tableConfigOperator.goneTableNotices(
          props.previousIdNameMap,
        ),
        tabsByTable: tabOfTable,
      }),
      ...this.rowIdsOperator.idFirstNotices(),
    ];
    const readings = this._traitReadings();
    const planned = this._withDeletionEffects(
      planner.plan({
        ...withApiAccessReadings({
          readings,
          previousIdNameMap: props.previousIdNameMap,
          liveTableIds: new Set(this.ss.raw.activeTableIds),
        }),
        sheetStops: [
          ...this.tableConfigOperator.grantHeadRowStops(accessChanges.grants),
          ...this.columnConfigOperator.headerTakenStops(planner.newColumns),
          ...this.conditionalFormatConfigOperator.noSuchRuleStops(
            planner.ruleRemovals,
            planner.ruleReplacements,
          ),
          ...this.conditionalFormatConfigOperator.unmodelableReplaceStops(
            planner.ruleReplacements,
          ),
        ],
      }),
      props.previousIdNameMap,
    );
    const isApplying = props.isSending && !planned.isStopped;
    // An unstopped dry run records the live readings: configTraits carry sends it won't make.
    this.columnConfigOperator.recordEmptyValueAllowed(
      props.isSending || planned.isStopped
        ? planned.configTraits
        : readings.live,
    );
    const newColumns = this.columnConfigOperator.mintColumnIds(
      planned.newColumns,
    );
    const ruleRemovals =
      this.conditionalFormatConfigOperator.indexedRuleRemovals(
        planned.ruleRemovals,
      );
    const applied: AppliedPlan = {
      accessChanges,
      sentTraits: planned.sentTraits,
      newColumns,
      columnMoves: this.columnConfigOperator.planColumnMoves(
        planned.sentTraits,
        newColumns,
      ),
      ruleRemovals,
      ruleReplacements:
        this.conditionalFormatConfigOperator.compiledRuleReplacements({
          replacements: planned.ruleReplacements,
          removals: ruleRemovals,
          sentTraits: planned.sentTraits,
          newColumns,
        }),
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
      const stopped = this._flushUnlessStopped(applied);
      // No files: column IDs already recorded would reach the configs but not the sheet.
      if (stopped !== undefined) {
        return {
          result: resultBuilder.flushStopped({
            planned,
            outcome: stopped,
            rowIds: { plan: checkboxPlan, notices },
          }),
        };
      }
    }
    const sent = isApplying ? planned.plan.map(({ subject }) => subject) : [];
    const previous: ConfigSnapshot = {
      idNameMap: props.previousIdNameMap,
      columnConfigs: columnConfigsByName(),
      valueConfigs: installedConfigs().valueConfigs,
    };
    let refetched: SyncCoordinator | undefined;
    let regeneration: Regeneration;
    // A fresh read, since Raw's state has no reset; only a sent sheet trait, a new column or a removed rule changes what it would read.
    if (
      isApplying &&
      (planned.sentTraits.some(isSheetTrait) ||
        applied.newColumns.length > 0 ||
        applied.ruleRemovals.length > 0 ||
        applied.ruleReplacements.length > 0)
    ) {
      refetched = SyncCoordinator.init();
      regeneration = refetched._refetchedRegeneration({
        configTraits: planned.configTraits,
        plan: planned,
        sent,
        previous,
        accessChanges,
        newColumns: applied.newColumns,
        ruleRemovals: applied.ruleRemovals,
        ruleReplacements: applied.ruleReplacements,
      });
    } else {
      regeneration = this._regeneration({ notes, sent, previous });
    }
    // The fill reads the configs the files came from.
    const regenerator = refetched ?? this;
    const { files, changes, didntLand, createdColumns } = regeneration;
    // After the files are strings, so the host's gate judges what a send would write and the configs never see it.
    if (props.isHostWithholding && !props.isSending && !planned.isStopped) {
      this._applyPlan(applied);
      this._flushPlan(applied);
    }
    const landed = sent.filter(hasLanded(didntLand));
    const nameRewriter = SyncNameRewriter.init({
      alterations: remainingAlterations(alterations, landed),
      configs: regeneration.configs,
      sent,
      landed,
    });
    if (landed.length > 0 || nameRewriter.rewrites.length > 0) {
      files.alterations = alterationsFileSource(nameRewriter.alterations);
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
        createdColumns,
        rowIds: { plan: [...checkboxPlan, ...fillPlan], notices },
        nameRewriter,
      }),
      files,
    };
  }
  // A pure planner can't read the rules a deleted Table takes with it, so the run fills them in from its live read.
  private _withDeletionEffects(
    planned: SyncPlannerResult,
    previousIdNameMap: IdNameMapGeneric,
  ): SyncPlannerResult {
    const tableIds = tableIdsByName(previousIdNameMap);
    const plan = planned.plan.map((item) => {
      const table = deletedTableOf(item.subject);
      if (table === undefined) return item;
      // The entry refuser and the planner already hold the Table to the map and the live sheet.
      const tableId = Val.assert(tableIds.get(table), `Table ID of ${table}`);
      const ruleEffects = this.conditionalFormatConfigOperator.deletionEffects({
        table,
        tableId,
      });
      return { ...item, effects: [...item.effects, ...ruleEffects] };
    });
    return { ...planned, plan };
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
    previous,
    accessChanges,
    newColumns,
    ruleRemovals,
    ruleReplacements,
  }: RefetchProps): Regeneration {
    // Recorded before the read, so its head-row fetch covers granted Tables and skips revoked ones.
    this.tableConfigOperator.recordApiAccessChanges(accessChanges);
    this._readLive();
    this.columnConfigOperator.recordEmptyValueAllowed(configTraits);
    const traitsDidntLand = didntLandOf({
      sentTraits: plan.sentTraits,
      plan: plan.plan,
      live: this._traitReadings().live,
    });
    const notes = this._recordUntypedColumns();
    const columns = this.columnConfigOperator.landedColumns(newColumns);
    const rulesDidntLand = [
      ...this.conditionalFormatConfigOperator.removalsDidntLand(ruleRemovals),
      ...this.conditionalFormatConfigOperator.replacementsDidntLand(
        ruleReplacements,
      ),
    ];
    return {
      ...this._regeneration({ notes, sent, previous }),
      didntLand: [...traitsDidntLand, ...columns.didntLand, ...rulesDidntLand],
      createdColumns: columns.createdColumns,
    };
  }
  private _regeneration({
    notes,
    sent,
    previous,
  }: RegenerationProps): Regeneration {
    this.valueConfigOperator.fetchAfterColumnConfigSynced();
    const configs: ConfigSnapshots = {
      previous,
      regenerated: {
        idNameMap: this._idNameMap(),
        columnConfigs: this.columnConfigOperator.newColumnConfigs(),
        valueConfigs: this.valueConfigOperator.newValueConfigs(),
      },
    };
    return {
      files: this._configFiles(configs.regenerated.idNameMap),
      configs,
      changes: typedTraitChanges({ ...configs, sent }),
      notes,
      didntLand: [],
      createdColumns: [],
    };
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
  private _applyPlan({
    accessChanges,
    sentTraits,
    newColumns,
    ruleRemovals,
    reads,
  }: AppliedPlan): void {
    this._readEnteredData(reads);
    // Indexed by the pre-send read, before any write here can stale the sheet's rule indexes.
    this.conditionalFormatConfigOperator.removeRules(ruleRemovals);
    // Before the access changes, so it sends what the pre-send read planned.
    this.rowIdsOperator.addCheckboxes();
    this._applyApiAccessChanges(accessChanges);
    this.columnConfigOperator
      .addMissingColumnIds()
      .appendColumns(newColumns)
      .updateHeaders(sentTraits)
      .updateFormulas(sentTraits);
    this.tableConfigOperator
      .updateTableNames(sentTraits)
      .deleteTables(sentTraits);
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
  // The Node host's gate refuses a held flush by throwing, as Google refuses a batch over part of a Table; any other error goes on up.
  private _flushUnlessStopped(
    applied: AppliedPlan,
  ): SyncFlushStoppedOutcome | undefined {
    try {
      this._flushPlan(applied);
      return undefined;
    } catch (error) {
      if (error instanceof HeldFlushRefusal) return heldStoppedOutcome(error);
      if (error instanceof PartialTableRefusal) {
        return refusedStoppedOutcome(
          this.ss.raw.landedFlushCount,
          error.message,
        );
      }
      throw error;
    }
  }
  // Column deletes and inserts with the replaced rules compiled against the layout they leave, then moves, each flush alone: every later write is indexed by the layout the one before leaves.
  private _flushPlan({
    sentTraits,
    newColumns,
    ruleReplacements,
    columnMoves,
  }: AppliedPlan): void {
    this.ss.batchUpdateGSheets();
    this.columnConfigOperator.updateDeletionsAndInserts(sentTraits, newColumns);
    this.conditionalFormatConfigOperator.replaceRules(ruleReplacements);
    this.ss.batchUpdateGSheets();
    this.columnConfigOperator.moveColumns(columnMoves);
    this.ss.batchUpdateGSheets();
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
  private _configFiles(idNameMap: IdNameMapGeneric): ConfigRegeneration {
    return {
      tableConfigs: this.tableConfigOperator.toFileSource(),
      tabConfigs: this.tabConfigOperator.toFileSource(),
      conditionalFormatConfigs:
        this.conditionalFormatConfigOperator.toFileSource(),
      retiredIdPrefixes: this.tableConfigOperator.retiredIdPrefixesFileSource(),
      columnConfigs: this.columnConfigOperator.toFileSource(),
      valueConfigs: this.valueConfigOperator.toFileSource(),
      idNameMap: idNameMapFileSource(idNameMap),
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

interface GoneNoticesProps {
  tabsGone: SyncTabGoneNotice[];
  tablesGone: SyncTableGoneNotice[];
  tabsByTable: Record<string, string | undefined>;
}

// Each gone tab's line comes just before its own Tables' lines, then the Tables gone from a live tab.
function goneNotices({
  tabsGone,
  tablesGone,
  tabsByTable,
}: GoneNoticesProps): SyncNotice[] {
  const goneTabs = new Set(tabsGone.map(({ tab }) => tab));
  const fromLiveTabs = tablesGone.filter(({ table }) => {
    const tab = tabsByTable[table];
    return tab === undefined || !goneTabs.has(tab);
  });
  return [
    ...tabsGone.flatMap((tabGone) => [
      tabGone,
      ...tablesGone.filter(({ table }) => tabsByTable[table] === tabGone.tab),
    ]),
    ...fromLiveTabs,
  ];
}
