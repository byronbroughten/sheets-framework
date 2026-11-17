import { Val } from "@byronbroughten/utils/val";

import type {
  BooleanCondition,
  ConditionalFormat,
  ConditionalFormatRule,
  ModelableConditionalFormatRule,
} from "../../00_Source/RawSource/ConditionalFormat";
import {
  type GridRangeProps,
  rangesShareCell,
} from "../../00_Source/RawSource/RawSource";
import {
  type HexColor,
  rgbChannelsHex,
} from "../../00_Source/RawSource/RgbColor";
import { SheetIndex } from "../../00_Source/RawSource/SheetIndex";
import { SpreadsheetBaseSchema } from "../SpreadsheetBaseSchema";
import { BraceFormula, type RuleLayout } from "./BraceFormula";
import type { SyncPlanEffect } from "./syncRunResult";

const schema = new SpreadsheetBaseSchema();

const handleSeparators = { part: " | ", column: "+" } as const;

interface HandleRename {
  handle: string;
  columnNow: (columnName: string) => string;
  formulaNow: (formula: string) => string;
}

export interface RuleRecordsTab {
  tabName: string;
  sheetGid: number;
  tables: RuleRecordsTable[];
  rules: ConditionalFormatRule[];
}

export interface RuleRecordsTable {
  tableKey: string;
  tableId: string;
  layout: RuleLayout;
}

export interface DeletedTableTab {
  tabName: string;
  rules: ConditionalFormatRule[];
  table: RuleRecordsTable;
}

export interface RuleRecordFormat {
  backgroundColor?: HexColor;
  foregroundColor?: HexColor;
}

export type RuleColumns = string[] | "all";

interface BooleanRuleContent {
  kind: "boolean";
  condition: BooleanCondition;
  format: RuleRecordFormat;
}

interface UnmodelableRuleContent {
  kind: "unmodelable";
  ranges: string[];
  rawContent: string;
}

interface TableRecordOwner {
  tableId: string;
}

interface TabRecordOwner {
  sheetGid: number;
}

export interface TableRuleRecord extends BooleanRuleContent, TableRecordOwner {
  columns: RuleColumns;
}

export interface TabRuleRecord extends BooleanRuleContent, TabRecordOwner {
  ranges: string[];
}

export interface UnmodelableTableRuleRecord
  extends UnmodelableRuleContent, TableRecordOwner {}

export interface UnmodelableTabRuleRecord
  extends UnmodelableRuleContent, TabRecordOwner {}

interface HandleParts {
  scope: string;
  condition: BooleanCondition;
  format: RuleRecordFormat;
}

interface BraceConditionProps {
  condition: BooleanCondition;
  table: RuleRecordsTable;
  isOneColumn: boolean;
}

export type TableRuleRecords = Record<
  string,
  TableRuleRecord | UnmodelableTableRuleRecord
>;
export type TabRuleRecords = Record<
  string,
  TabRuleRecord | UnmodelableTabRuleRecord
>;

export interface RuleRecordSet {
  tables: Record<string, TableRuleRecords>;
  tabs: Record<string, TabRuleRecords>;
}

interface HandledRecord<RR> {
  handle: string;
  record: RR;
}

// A Table rule's place in its tab's rule list; `sharedHandle` is its handle before any duplicate suffix, which every copy of the rule has.
export interface TableRuleIndex {
  sharedHandle: string;
  index: number;
  kind: TableRuleRecords[string]["kind"];
}

// A Table's modelled rule ending above the Table's last body row; `endRowIndex` is its ranges' exclusive end.
export interface ShortTableRule {
  handle: string;
  index: number;
  record: TableRuleRecord;
  endRowIndex: number;
}

// A rule a Table owns, with its place in its tab's rule list.
export interface OwnedTableRule {
  handle: string;
  index: number;
  rule: ConditionalFormatRule;
  record: TableRuleRecord | UnmodelableTableRuleRecord;
}

export interface RuleText {
  range: string;
  formula: string;
}

interface FiledRuleBase<RO, RR> {
  owner: RO;
  handle: string;
  sharedHandle: string;
  record: RR;
  index: number;
}

type FiledRule =
  | FiledRuleBase<
      RuleRecordsTable,
      TableRuleRecord | UnmodelableTableRuleRecord
    >
  | FiledRuleBase<undefined, TabRuleRecord | UnmodelableTabRuleRecord>;

// A handle is readable content, not a hash: packages/utils has no hash function.
export const RuleRecords = {
  build(tabs: RuleRecordsTab[]): RuleRecordSet {
    return tabs.reduce<RuleRecordSet>(
      (recordSet, tab) => {
        for (const filed of filedRules(tab)) {
          if (filed.owner === undefined) {
            (recordSet.tabs[tab.tabName] ??= {})[filed.handle] = filed.record;
            continue;
          }
          const group = (recordSet.tables[filed.owner.tableKey] ??= {});
          group[filed.handle] = filed.record;
        }
        return recordSet;
      },
      { tables: {}, tabs: {} },
    );
  },
  // Filed by build's walk, so a duplicate's suffix names the same rule here as in the configs.
  ruleIndexesByHandle(
    tab: RuleRecordsTab,
    tableId: string,
  ): Map<string, TableRuleIndex> {
    return new Map(
      filedRules(tab).flatMap(
        ({
          owner,
          handle,
          sharedHandle,
          index,
          record,
        }): [string, TableRuleIndex][] => {
          if (owner?.tableId !== tableId) return [];
          return [[handle, { sharedHandle, index, kind: record.kind }]];
        },
      ),
    );
  },
  // Filed by build's walk, so each handle names the rule the configs record.
  tableRules(tab: RuleRecordsTab, table: RuleRecordsTable): OwnedTableRule[] {
    return filedRules(tab).flatMap((filed): OwnedTableRule[] => {
      if (filed.owner === undefined) return [];
      if (filed.owner.tableId !== table.tableId) return [];
      const { handle, record, index } = filed;
      const rule = Val.assert(tab.rules[index], `rule ${index}`);
      return [{ handle, index, rule, record }];
    });
  },
  // A rule with no brace form isn't one: sync can't compile it against the longer span.
  shortRules(tab: RuleRecordsTab, table: RuleRecordsTable): ShortTableRule[] {
    const owned = RuleRecords.tableRules(tab, table);
    return owned.flatMap(
      ({ handle, index, rule, record }): ShortTableRule[] => {
        if (record.kind !== "boolean") return [];
        const endRowIndex = Math.max(
          ...rule.ranges.map((range) =>
            Val.assert(range.endRowIndex, `end row of ${handle}`),
          ),
        );
        if (endRowIndex >= table.layout.endRowIndex) return [];
        return [{ handle, index, record, endRowIndex }];
      },
    );
  },
  // A rule as a check shows it: its ranges in A1 and its formula, or its raw content when it has none.
  ruleText(rule: ConditionalFormatRule): RuleText {
    return {
      range: rule.ranges.map((range) => schema.gridRangeA1(range)).join(","),
      formula: ruleFormula(rule),
    };
  },
  // The handle a rule would have on the Table before any duplicate suffix.
  sharedHandleOf(rule: ConditionalFormatRule, table: RuleRecordsTable): string {
    return tableRecord(rule, table).handle;
  },
  // An unmodelable rule's handle comes back unchanged: its scope is ranges, and its condition no formula.
  renamedHandle({ handle, columnNow, formulaNow }: HandleRename): string {
    const { part, column } = handleSeparators;
    const parts = handle.split(part);
    const [scope] = parts;
    const format = parts.at(-1);
    if (parts.length < 3 || scope === undefined || format === undefined) {
      return handle;
    }
    const condition = parts.slice(1, -1).join(part);
    let scopeNow = scope;
    if (scope !== "all") {
      scopeNow = scope
        .split(column)
        .map((columnName) => columnNow(columnName))
        .join(column);
    }
    let conditionNow = condition;
    if (condition.startsWith("=")) conditionNow = formulaNow(condition);
    return [scopeNow, conditionNow, format].join(part);
  },
  // Nothing is sent for a partly-over rule: it stays on the sheet, so the plan only reports it.
  deletionEffects({
    tabName,
    rules,
    table,
  }: DeletedTableTab): SyncPlanEffect[] {
    const removed: TableRuleRecords = {};
    const partlyOver: SyncPlanEffect[] = [];
    for (const rule of rules) {
      if (!rule.ranges.some((range) => touchesTable(range, table))) continue;
      if (rule.ranges.every((range) => isInsideTable(range, table))) {
        addRecord(removed, tableRecord(rule, table));
        continue;
      }
      partlyOver.push({
        kind: "rulePartlyOver",
        tab: tabName,
        ...RuleRecords.ruleText(rule),
      });
    }
    const handles = Object.keys(removed);
    if (handles.length === 0) return partlyOver;
    return [{ kind: "rulesRemoved", handles }, ...partlyOver];
  },
};

// The first of the tab's Tables the rule touches, as build files it.
function ownerTable(
  rule: ConditionalFormatRule,
  tab: RuleRecordsTab,
): RuleRecordsTable | undefined {
  return tab.tables.find((candidate) =>
    rule.ranges.some((range) => touchesTable(range, candidate)),
  );
}

function touchesTable(range: GridRangeProps, table: RuleRecordsTable): boolean {
  const { startCol, endCol } = tableColumnBounds(table);
  return rangesShareCell(range, {
    sheetId: range.sheetId,
    startRowIndex: table.layout.headerTopLeft.rowIndex,
    endRowIndex: table.layout.endRowIndex,
    startColumnIndex: SheetIndex.col(startCol),
    endColumnIndex: SheetIndex.col(endCol),
  });
}

function tableColumnBounds(table: RuleRecordsTable): {
  startCol: number;
  endCol: number;
} {
  const startCol = table.layout.headerTopLeft.colIndex;
  return {
    startCol,
    endCol: startCol + table.layout.columnNamesLeftToRight.length,
  };
}

// In the tab's rule order, each owner's duplicates suffixed as they come.
function filedRules(tab: RuleRecordsTab): FiledRule[] {
  const takenByOwner = new Map<RuleRecordsTable | undefined, Set<string>>();
  return tab.rules.map((rule, index): FiledRule => {
    const owner = ownerTable(rule, tab);
    const taken = takenByOwner.get(owner) ?? new Set<string>();
    takenByOwner.set(owner, taken);
    if (owner === undefined) {
      const { handle, record } = tabRecord(rule, tab.sheetGid);
      const suffixed = takeHandle(taken, handle);
      return { owner, handle: suffixed, sharedHandle: handle, record, index };
    }
    const { handle, record } = tableRecord(rule, owner);
    const suffixed = takeHandle(taken, handle);
    return { owner, handle: suffixed, sharedHandle: handle, record, index };
  });
}

function takeHandle(taken: Set<string>, handle: string): string {
  const suffixed = unusedHandle(handle, (candidate) => taken.has(candidate));
  taken.add(suffixed);
  return suffixed;
}

function addRecord<RR>(
  group: Record<string, RR>,
  { handle, record }: HandledRecord<RR>,
): void {
  group[unusedHandle(handle, (candidate) => group[candidate] !== undefined)] =
    record;
}

function unusedHandle(
  handle: string,
  isTaken: (candidate: string) => boolean,
): string {
  let suffixed = handle;
  let repeat = 1;
  while (isTaken(suffixed)) {
    repeat += 1;
    suffixed = `${handle}-${repeat}`;
  }
  return suffixed;
}

function tabRecord(
  rule: ConditionalFormatRule,
  sheetGid: number,
): HandledRecord<TabRuleRecord | UnmodelableTabRuleRecord> {
  const ranges = rule.ranges.map((range) => schema.gridRangeA1(range));
  if (rule.kind === "unmodelable") {
    const { rawContent } = rule;
    return {
      handle: unmodelableHandle(ranges, rawContent),
      record: { kind: "unmodelable", sheetGid, ranges, rawContent },
    };
  }
  const format = recordFormat(rule.format);
  return {
    handle: handleOf({
      scope: ranges.join(","),
      condition: rule.condition,
      format,
    }),
    record: {
      kind: "boolean",
      sheetGid,
      ranges,
      condition: rule.condition,
      format,
    },
  };
}

function tableRecord(
  rule: ConditionalFormatRule,
  table: RuleRecordsTable,
): HandledRecord<TableRuleRecord | UnmodelableTableRuleRecord> {
  const { tableId } = table;
  const ranges = rule.ranges.map((range) => schema.gridRangeA1(range));
  const modelled =
    rule.kind === "boolean" ? modelledRecord(rule, table) : undefined;
  if (modelled !== undefined) return modelled;
  const rawContent = rawContentOf(rule);
  return {
    handle: unmodelableHandle(ranges, rawContent),
    record: { kind: "unmodelable", tableId, ranges, rawContent },
  };
}

// Undefined when a range leaves the Table's body or the formula has no brace form.
function modelledRecord(
  rule: ModelableConditionalFormatRule,
  table: RuleRecordsTable,
): HandledRecord<TableRuleRecord> | undefined {
  const columnNames = coveredColumnNames(rule.ranges, table);
  if (columnNames === undefined) return undefined;
  const condition = braceCondition({
    condition: rule.condition,
    table,
    isOneColumn: columnNames.length === 1,
  });
  if (condition === undefined) return undefined;
  let columns: RuleColumns = columnNames;
  if (columnNames.length === table.layout.columnNamesLeftToRight.length) {
    columns = "all";
  }
  const format = recordFormat(rule.format);
  return {
    handle: handleOf({ scope: columnsText(columns), condition, format }),
    record: {
      kind: "boolean",
      tableId: table.tableId,
      columns,
      condition,
      format,
    },
  };
}

function coveredColumnNames(
  ranges: GridRangeProps[],
  table: RuleRecordsTable,
): string[] | undefined {
  const offsets = new Set<number>();
  for (const range of ranges) {
    const rangeOffsets = bodyColumnOffsets(range, table);
    if (rangeOffsets === undefined) return undefined;
    rangeOffsets.forEach((offset) => offsets.add(offset));
  }
  const columnNames = table.layout.columnNamesLeftToRight.filter((_, offset) =>
    offsets.has(offset),
  );
  if (columnNames.includes(undefined)) return undefined;
  return columnNames.filter((columnName) => columnName !== undefined);
}

// Undefined unless the range starts on the first body row and stays within the Table.
function bodyColumnOffsets(
  {
    startRowIndex,
    endRowIndex,
    startColumnIndex,
    endColumnIndex,
  }: GridRangeProps,
  table: RuleRecordsTable,
): number[] | undefined {
  const { startCol, endCol } = tableColumnBounds(table);
  if (startRowIndex !== table.layout.headerTopLeft.rowIndex + 1) {
    return undefined;
  }
  if (endRowIndex === undefined || endRowIndex > table.layout.endRowIndex) {
    return undefined;
  }
  if (startColumnIndex === undefined || startColumnIndex < startCol) {
    return undefined;
  }
  if (endColumnIndex === undefined || endColumnIndex > endCol) {
    return undefined;
  }
  return Array.from(
    { length: endColumnIndex - startColumnIndex },
    (_, step) => startColumnIndex - startCol + step,
  );
}

function braceCondition({
  condition,
  table,
  isOneColumn,
}: BraceConditionProps): BooleanCondition | undefined {
  if (condition.type !== "CUSTOM_FORMULA") return condition;
  const formula = BraceFormula.decompileForRule({
    layout: table.layout,
    formula: condition.formula,
    isOneColumn,
  });
  if (formula === undefined) return undefined;
  return { type: "CUSTOM_FORMULA", formula };
}

function columnsText(columns: RuleColumns): string {
  if (columns === "all") return "all";
  return columns.join(handleSeparators.column);
}

function handleOf({ scope, condition, format }: HandleParts): string {
  return [scope, conditionText(condition), formatText(format)].join(
    handleSeparators.part,
  );
}

function conditionText(condition: BooleanCondition): string {
  if (condition.type === "CUSTOM_FORMULA") return condition.formula;
  return `${condition.type} ${JSON.stringify(condition.value)}`;
}

function formatText({
  backgroundColor,
  foregroundColor,
}: RuleRecordFormat): string {
  const parts: string[] = [];
  if (backgroundColor !== undefined) {
    parts.push(`background ${backgroundColor}`);
  }
  if (foregroundColor !== undefined) parts.push(`text ${foregroundColor}`);
  return parts.join(", ");
}

function unmodelableHandle(ranges: string[], rawContent: string): string {
  return `${ranges.join(",")} | unmodelable ${rawContent}`;
}

// A boolean rule that doesn't fit its Table keeps its condition and format as its raw content.
function rawContentOf(rule: ConditionalFormatRule): string {
  if (rule.kind === "unmodelable") return rule.rawContent;
  return JSON.stringify({ condition: rule.condition, format: rule.format });
}

function recordFormat({
  backgroundColor,
  foregroundColor,
}: ConditionalFormat): RuleRecordFormat {
  return {
    ...(backgroundColor !== undefined
      ? { backgroundColor: rgbChannelsHex(backgroundColor) }
      : {}),
    ...(foregroundColor !== undefined
      ? { foregroundColor: rgbChannelsHex(foregroundColor) }
      : {}),
  };
}

// The fake's deleteTable replay drops a rule by the same test: an open bound reaches past the Table.
function isInsideTable(
  {
    startRowIndex,
    endRowIndex,
    startColumnIndex,
    endColumnIndex,
  }: GridRangeProps,
  table: RuleRecordsTable,
): boolean {
  const { startCol, endCol } = tableColumnBounds(table);
  if (endRowIndex === undefined || endColumnIndex === undefined) return false;
  if (startColumnIndex === undefined) return false;
  return (
    table.layout.headerTopLeft.rowIndex <= startRowIndex &&
    endRowIndex <= table.layout.endRowIndex &&
    startCol <= startColumnIndex &&
    endColumnIndex <= endCol
  );
}

function ruleFormula(rule: ConditionalFormatRule): string {
  if (rule.kind === "unmodelable") return rule.rawContent;
  return conditionText(rule.condition);
}
