import { Val } from "@byronbroughten/utils/val";

import type { ModelableConditionalFormatRule } from "../00_Source/RawSource/ConditionalFormat";
import type { RuleLayout } from "../01_SpreadsheetSchema/configReaders/BraceFormula";
import { columnConfigsByName } from "../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import { IntendedRule } from "../01_SpreadsheetSchema/configReaders/IntendedRule";
import {
  type ConditionalFormatConfigsBase,
  makeImportLine,
} from "../01_SpreadsheetSchema/configReaders/makeConfigs";
import {
  RuleRecords,
  type RuleRecordsTab,
  type RuleRecordsTable,
  type TableRuleIndex,
} from "../01_SpreadsheetSchema/configReaders/RuleRecords";
import {
  type IntendedTrait,
  type PlannedRuleAddition,
  type PlannedRuleRemoval,
  type PlannedRuleReplacement,
  ruleAdditionLines,
  ruleRemovalLines,
  ruleReplacementLines,
} from "../01_SpreadsheetSchema/configReaders/SyncPlanner";
import type { WideRule } from "../01_SpreadsheetSchema/configReaders/SyncPlanner/wideAlterations";
import type {
  SyncNoSuchRuleStop,
  SyncPlanEffect,
  SyncSendLine,
  SyncUnmodelableRuleStop,
} from "../01_SpreadsheetSchema/configReaders/syncRunResult";
import type { SheetRaw } from "../02_SpreadsheetRaw/SheetRaw";
import type { TableRaw } from "../02_SpreadsheetRaw/TableRaw";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import {
  type ApiTableColumns,
  ColumnConfigsOperator,
  type MintedNewColumn,
} from "./ColumnConfigsOperator";
import {
  columnNamesByColumnId,
  oneLinePerLeafFileSource,
  tablesAndTabsFileSource,
} from "./configFileSource";
import { SpreadsheetBaseOperator } from "./SpreadsheetBaseOperator";
import { TableConfigsOperator } from "./TableConfigsOperator";

// A Table a plan deletes: its codebase name and its ID.
export interface DeletedTable {
  table: string;
  tableId: string;
}

// A removal matched on the pre-send read; `copiesLeft` is how many rules with its shared handle the send leaves on the Table.
export interface IndexedRuleRemoval
  extends PlannedRuleRemoval, Omit<TableRuleIndex, "kind"> {
  tableId: string;
  copiesLeft: number;
}

// A replacement matched on the pre-send read, `index` already moved past the removals sent before it, and compiled against the planned layout.
export interface CompiledRuleReplacement extends PlannedRuleReplacement {
  tableId: string;
  sheetGid: number;
  index: number;
  compiled: ModelableConditionalFormatRule;
}

// The run's column deletes and inserts, which move the layout a rule compiles against.
export interface PlannedColumnChanges {
  sentTraits: IntendedTrait[];
  newColumns: MintedNewColumn[];
}

export interface CompiledRuleReplacementsProps {
  replacements: PlannedRuleReplacement[];
  removals: IndexedRuleRemoval[];
  additions: PlannedRuleAddition[];
  columnChanges: PlannedColumnChanges;
}

// An add compiled against the planned layout; `copiesAfter` is how many rules with its shared handle the Table holds once the run's sends land.
export interface CompiledRuleAddition extends PlannedRuleAddition {
  sheetGid: number;
  compiled: ModelableConditionalFormatRule;
  copiesAfter: number;
}

export interface CompiledRuleAdditionsProps {
  additions: PlannedRuleAddition[];
  removals: IndexedRuleRemoval[];
  replacements: CompiledRuleReplacement[];
  columnChanges: PlannedColumnChanges;
}

interface CompiledOnPlanProps {
  tableKey: string;
  rules: TableRules;
  rule: WideRule;
  columnChanges: PlannedColumnChanges;
}

interface CompiledOnPlan {
  compiled: ModelableConditionalFormatRule;
  plannedTable: RuleRecordsTable;
}

interface PlannedLayoutOfProps {
  table: string;
  layout: RuleLayout;
  columnChanges: PlannedColumnChanges;
}

interface CopiesKeptProps {
  table: string;
  sharedHandle: string;
  rules: TableRules;
  plannedTable: RuleRecordsTable;
  removals: IndexedRuleRemoval[];
  replacements: CompiledRuleReplacement[];
}

interface TableRules {
  tableId: string;
  sheetGid: number;
  table: RuleRecordsTable;
  indexes: Map<string, TableRuleIndex>;
}

interface HandleEntry {
  table: string;
  handle: string;
  path: string;
}

// Reads the rules the regen's plain property get brought into Raw, so it costs no round trip of its own.
export class ConditionalFormatConfigOperator extends SpreadsheetBaseOperator {
  static init(): ConditionalFormatConfigOperator {
    return new ConditionalFormatConfigOperator(
      SpreadsheetBaseOperator.initOperatorProps(),
    );
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
  newConditionalFormatConfigs(): ConditionalFormatConfigsBase {
    const apiTableColumnsByTableId =
      this.columnConfigOperator.apiTableColumnsByTableId();
    const { raw } = this.ss;
    const tabs = raw.activeSheetGids.map((sheetGid) =>
      this._ruleRecordsTab(raw.sheet(sheetGid), apiTableColumnsByTableId),
    );
    return RuleRecords.build(tabs);
  }
  private _ruleRecordsTab(
    sheet: SheetRaw,
    apiTableColumnsByTableId: Map<string, ApiTableColumns>,
  ): RuleRecordsTab {
    const tables = sheet.tableIds.flatMap((tableId) => {
      const apiTableColumns = apiTableColumnsByTableId.get(tableId);
      if (apiTableColumns === undefined) return [];
      return [this._ruleRecordsTable(tableId, apiTableColumns)];
    });
    return {
      tabName: this.schema.titleToName(sheet.title),
      sheetGid: sheet.sheetGid,
      tables,
      rules: sheet.conditionalFormatRules(),
    };
  }
  // Named from the previous configs: apiTableColumnsByTableId holds only after the regen, which no longer has the Table.
  deletionEffects({
    table: tableKey,
    tableId,
  }: DeletedTable): SyncPlanEffect[] {
    const columnConfigs = Val.assert(
      columnConfigsByName()[tableKey],
      `previous column configs of ${tableKey}`,
    );
    const { sheet } = this.ss.raw.table(tableId);
    return RuleRecords.deletionEffects({
      tabName: this.schema.titleToName(sheet.title),
      rules: sheet.conditionalFormatRules(),
      table: this._ruleRecordsTable(tableId, { tableKey, columnConfigs }),
    });
  }
  private _ruleRecordsTable(
    tableId: string,
    { tableKey, columnConfigs }: ApiTableColumns,
  ): RuleRecordsTable {
    const table = this.ss.raw.table(tableId);
    const columnNamesById = columnNamesByColumnId(columnConfigs);
    const columnNamesLeftToRight = table.fullTableColIndexes.map((colIndex) =>
      columnNameAt({ table, colIndex, columnNamesById }),
    );
    return {
      tableKey,
      tableId,
      layout: {
        headerTopLeft: {
          rowIndex: table.startRowIndex,
          colIndex: table.startColumnIndex,
        },
        endRowIndex: table.endRowIndex,
        columnNamesLeftToRight,
      },
    };
  }
  // A handle a person changed since the last sync names no rule now; deleting by a guess could take the wrong one.
  noSuchRuleStops(
    removals: PlannedRuleRemoval[],
    replacements: PlannedRuleReplacement[],
  ): SyncNoSuchRuleStop[] {
    const entries = handleEntries(removals, replacements);
    const rulesByTable = this._tableRulesByTable(entries);
    return entries.flatMap(({ table, handle, path }): SyncNoSuchRuleStop[] => {
      // A Table a person deleted took its rules with it, so the planner drops the entry.
      const rules = rulesByTable.get(table);
      if (rules === undefined || rules.indexes.has(handle)) return [];
      return [
        { kind: "noSuchRule", table, path, handles: [...rules.indexes.keys()] },
      ];
    });
  }
  // Sync can't write a rule it can't model, so a replace would lose whatever the rule held.
  unmodelableReplaceStops(
    replacements: PlannedRuleReplacement[],
  ): SyncUnmodelableRuleStop[] {
    const rulesByTable = this._tableRulesByTable(replacements);
    return replacements.flatMap((replacement): SyncUnmodelableRuleStop[] => {
      const rule = rulesByTable
        .get(replacement.table)
        ?.indexes.get(replacement.handle);
      if (rule?.kind !== "unmodelable") return [];
      return [
        {
          kind: "unmodelableRule",
          table: replacement.table,
          path: ruleReplacementLines.path(replacement),
          handle: replacement.handle,
        },
      ];
    });
  }
  // Compiled before any send, so a bad colour or brace throws with nothing sent; indexed for the flush after the removals', past the adds sent ahead of it.
  compiledRuleReplacements({
    replacements,
    removals,
    additions,
    columnChanges,
  }: CompiledRuleReplacementsProps): CompiledRuleReplacement[] {
    const rulesByTable = this._tableRulesByTable([
      ...replacements,
      ...additions,
    ]);
    return replacements.map((replacement) => {
      const rules = Val.assert(
        rulesByTable.get(replacement.table),
        `rules of ${replacement.table}`,
      );
      const { tableId, sheetGid, indexes } = rules;
      const { index } = Val.assert(
        indexes.get(replacement.handle),
        `rule ${replacement.handle} on ${replacement.table}`,
      );
      const removedAbove = removals.filter(
        (removal) =>
          removal.index < index &&
          this.ss.raw.table(removal.tableId).sheetGid === sheetGid,
      ).length;
      // The flusher sends a batch's adds before its updates, and each add goes in at index 0.
      const addedAbove = additions.filter(
        (addition) => rulesByTable.get(addition.table)?.sheetGid === sheetGid,
      ).length;
      const { compiled } = compiledOnPlan({
        tableKey: replacement.table,
        rules,
        rule: replacement.rule,
        columnChanges,
      });
      return {
        ...replacement,
        tableId,
        sheetGid,
        index: index - removedAbove + addedAbove,
        compiled,
      };
    });
  }
  // Compiled before any send, so a bad colour or brace throws with nothing sent.
  compiledRuleAdditions({
    additions,
    removals,
    replacements,
    columnChanges,
  }: CompiledRuleAdditionsProps): CompiledRuleAddition[] {
    const rulesByTable = this._tableRulesByTable(additions);
    const compiledAdditions = additions.map((addition) => {
      const rules = Val.assert(
        rulesByTable.get(addition.table),
        `rules of ${addition.table}`,
      );
      const { compiled, plannedTable } = compiledOnPlan({
        tableKey: addition.table,
        rules,
        rule: addition.rule,
        columnChanges,
      });
      // Named on the planned layout, as the refetch names it: a column the run deletes would move a read-layout name.
      const sharedHandle = RuleRecords.sharedHandleOf(compiled, plannedTable);
      return {
        addition,
        sheetGid: rules.sheetGid,
        compiled,
        sharedHandle,
        copiesKept: copiesKeptOf({
          table: addition.table,
          sharedHandle,
          rules,
          plannedTable,
          removals,
          replacements,
        }),
      };
    });
    return compiledAdditions.map(
      ({ addition, sheetGid, compiled, sharedHandle, copiesKept }) => {
        const addedCopies = compiledAdditions.filter(
          (other) =>
            other.addition.table === addition.table &&
            other.sharedHandle === sharedHandle,
        ).length;
        return {
          ...addition,
          sheetGid,
          compiled,
          copiesAfter: copiesKept + addedCopies,
        };
      },
    );
  }
  // Each at index 0, last first, so the list's first rule ends first and every rule already there stays below them.
  addRules(additions: CompiledRuleAddition[]): this {
    for (const { sheetGid, compiled } of [...additions].reverse()) {
      this.ss.raw.sheet(sheetGid).addConditionalFormatRuleFirst(compiled);
    }
    return this;
  }
  // By copies on the refetched layout: an add equal to a rule already there gives it a second copy, not a new handle.
  additionsDidntLand(additions: CompiledRuleAddition[]): SyncSendLine[] {
    const rulesByTable = this._tableRulesByTable(additions);
    return additions.flatMap((addition): SyncSendLine[] => {
      if (refetchedCopyCount(rulesByTable, addition) >= addition.copiesAfter) {
        return [];
      }
      const { subject, what } = ruleAdditionLines.sendLine(addition);
      return [{ subject, what: `${what}; Sheets didn't add it` }];
    });
  }
  // Queued for the flush after the removals', so the index skips the stale check the removals' flush set.
  replaceRules(replacements: CompiledRuleReplacement[]): this {
    for (const { sheetGid, index, compiled } of replacements) {
      this.ss.raw
        .sheet(sheetGid)
        .updateConditionalFormatRuleReindexed(index, compiled);
    }
    return this;
  }
  // By the compiled rule's handle on the refetched layout: compile then read back needn't give the intended rule's text.
  replacementsDidntLand(
    replacements: CompiledRuleReplacement[],
  ): SyncSendLine[] {
    const rulesByTable = this._tableRulesByTable(replacements);
    return replacements.flatMap((replacement): SyncSendLine[] => {
      if (refetchedCopyCount(rulesByTable, replacement) > 0) return [];
      const { subject, what } = ruleReplacementLines.sendLine(replacement);
      return [{ subject, what: `${what}; Sheets kept the old one` }];
    });
  }
  // Indexed on the read the stops were matched on, so every handle names a rule.
  indexedRuleRemovals(removals: PlannedRuleRemoval[]): IndexedRuleRemoval[] {
    const rulesByTable = this._tableRulesByTable(removals);
    return removals.map((removal) => {
      const { tableId, indexes } = Val.assert(
        rulesByTable.get(removal.table),
        `rules of ${removal.table}`,
      );
      const { sharedHandle, index } = Val.assert(
        indexes.get(removal.handle),
        `rule ${removal.handle} on ${removal.table}`,
      );
      const removedCopies = removals.filter(
        (other) =>
          other.table === removal.table &&
          indexes.get(other.handle)?.sharedHandle === sharedHandle,
      ).length;
      return {
        ...removal,
        tableId,
        index,
        sharedHandle,
        copiesLeft: copyCount(indexes, sharedHandle) - removedCopies,
      };
    });
  }
  removeRules(removals: IndexedRuleRemoval[]): this {
    for (const { tableId, index } of removals) {
      this.ss.raw
        .table(tableId)
        .sheet.removeConditionalFormatRuleByIndex(index);
    }
    return this;
  }
  // By copies, not handle: removing one of two equal rules renames the other to the removed one's handle.
  removalsDidntLand(removals: IndexedRuleRemoval[]): SyncSendLine[] {
    const rulesByTable = this._tableRulesByTable(removals);
    return removals.flatMap((removal): SyncSendLine[] => {
      const indexes = rulesByTable.get(removal.table)?.indexes ?? new Map();
      if (copyCount(indexes, removal.sharedHandle) <= removal.copiesLeft) {
        return [];
      }
      const { subject, what } = ruleRemovalLines.sendLine(removal);
      return [{ subject, what: `${what}; Sheets kept it` }];
    });
  }
  // Named as the regen names them, so a handle matches the one the configs record; one tab walk per Table.
  private _tableRulesByTable(
    entries: Pick<PlannedRuleRemoval, "table">[],
  ): Map<string, TableRules> {
    const rulesByTable = new Map<string, TableRules>();
    if (entries.length === 0) return rulesByTable;
    const apiTableColumnsByTableId =
      this.columnConfigOperator.identifiedApiTableColumnsByTableId();
    for (const table of new Set(entries.map((entry) => entry.table))) {
      const tableId = this.tableConfigOperator.tableIdOf(table);
      if (!this.ss.raw.tableIdIsActive(tableId)) continue;
      const tab = this._ruleRecordsTab(
        this.ss.raw.table(tableId).sheet,
        apiTableColumnsByTableId,
      );
      rulesByTable.set(table, {
        tableId,
        sheetGid: tab.sheetGid,
        table: Val.assert(
          tab.tables.find((candidate) => candidate.tableId === tableId),
          `rule layout of ${table}`,
        ),
        indexes: RuleRecords.ruleIndexesByHandle(tab, tableId),
      });
    }
    return rulesByTable;
  }
  toFileSource(): string {
    return [
      makeImportLine(["makeConditionalFormatConfigs"]),
      ``,
      `export const conditionalFormatConfigs = makeConditionalFormatConfigs(${tablesAndTabsFileSource(
        this.newConditionalFormatConfigs(),
        oneLinePerLeafFileSource,
      )});`,
      ``,
    ].join("\n");
  }
}

function compiledOnPlan({
  tableKey,
  rules,
  rule,
  columnChanges,
}: CompiledOnPlanProps): CompiledOnPlan {
  const plannedTable: RuleRecordsTable = {
    ...rules.table,
    layout: plannedLayoutOf({
      table: tableKey,
      layout: rules.table.layout,
      columnChanges,
    }),
  };
  return {
    compiled: IntendedRule.compile({
      tableKey,
      sheetGid: rules.sheetGid,
      layout: plannedTable.layout,
      rule,
    }),
    plannedTable,
  };
}

function plannedLayoutOf({
  table,
  layout,
  columnChanges: { sentTraits, newColumns },
}: PlannedLayoutOfProps): RuleLayout {
  return IntendedRule.plannedLayout({
    layout,
    deletedColumns: sentTraits.flatMap((intended) => {
      if (intended.kind !== "column" || intended.trait !== "deleted") {
        return [];
      }
      if (intended.table !== table) return [];
      return [intended.column];
    }),
    newColumns: newColumns.filter((newColumn) => newColumn.table === table),
  });
}

// The read's copies, less those the run removes or replaces away, plus the replacements that compile to it.
function copiesKeptOf({
  table,
  sharedHandle,
  rules,
  plannedTable,
  removals,
  replacements,
}: CopiesKeptProps): number {
  const removedCopies = removals.filter(
    (removal) =>
      removal.table === table && removal.sharedHandle === sharedHandle,
  ).length;
  const tableReplacements = replacements.filter(
    (replacement) => replacement.table === table,
  );
  const replacedAway = tableReplacements.filter(
    (replacement) =>
      rules.indexes.get(replacement.handle)?.sharedHandle === sharedHandle,
  ).length;
  const replacedInto = tableReplacements.filter(
    (replacement) =>
      RuleRecords.sharedHandleOf(replacement.compiled, plannedTable) ===
      sharedHandle,
  ).length;
  return (
    copyCount(rules.indexes, sharedHandle) -
    removedCopies -
    replacedAway +
    replacedInto
  );
}

// Named on the refetched layout; 0 for a Table the refetch no longer has.
function refetchedCopyCount(
  rulesByTable: Map<string, TableRules>,
  { table, compiled }: Pick<CompiledRuleAddition, "table" | "compiled">,
): number {
  const rules = rulesByTable.get(table);
  if (rules === undefined) return 0;
  return copyCount(
    rules.indexes,
    RuleRecords.sharedHandleOf(compiled, rules.table),
  );
}

function handleEntries(
  removals: PlannedRuleRemoval[],
  replacements: PlannedRuleReplacement[],
): HandleEntry[] {
  return [
    ...removals.map(({ table, handle }) => ({
      table,
      handle,
      path: ruleRemovalLines.path({ table, handle }),
    })),
    ...replacements.map(({ table, handle }) => ({
      table,
      handle,
      path: ruleReplacementLines.path({ handle }),
    })),
  ];
}

function copyCount(
  indexes: Map<string, TableRuleIndex>,
  sharedHandle: string,
): number {
  return [...indexes.values()].filter(
    (rule) => rule.sharedHandle === sharedHandle,
  ).length;
}

interface ColumnNameAtProps {
  table: TableRaw;
  colIndex: number;
  columnNamesById: Map<string, string>;
}

// Undefined for a column with no ID: a run that mints none leaves it out of the configs.
function columnNameAt({
  table,
  colIndex,
  columnNamesById,
}: ColumnNameAtProps): string | undefined {
  const columnId = table.columnResolver.columnIdAt(colIndex);
  if (columnId === "") return undefined;
  return Val.assert(
    columnNamesById.get(columnId),
    `column name at column index ${colIndex} of ${table.name}`,
  );
}
