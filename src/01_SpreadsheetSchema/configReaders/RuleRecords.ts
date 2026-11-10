import type {
  BooleanCondition,
  ConditionalFormat,
  ConditionalFormatRule,
  ModelableConditionalFormatRule,
} from "../../00_Source/RawSource/ConditionalFormat";
import type { GridRangeProps } from "../../00_Source/RawSource/RawSource";
import { rgbChannelsHex } from "../../00_Source/RawSource/RgbColor";
import {
  SheetIndex,
  type SheetRowIndex,
} from "../../00_Source/RawSource/SheetIndex";
import { SpreadsheetBaseSchema } from "../SpreadsheetBaseSchema";
import { BraceFormula, type RuleLayout } from "./BraceFormula";

const schema = new SpreadsheetBaseSchema();

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
  endRowIndex: SheetRowIndex;
}

export interface RuleRecordFormat {
  backgroundColor?: `#${string}`;
  foregroundColor?: `#${string}`;
}

export interface TableRuleRecord {
  kind: "boolean";
  tableId: string;
  columns: string[] | "all";
  condition: BooleanCondition;
  format: RuleRecordFormat;
}

export interface TabRuleRecord {
  kind: "boolean";
  sheetGid: number;
  ranges: string[];
  condition: BooleanCondition;
  format: RuleRecordFormat;
}

export interface UnmodelableTableRuleRecord {
  kind: "unmodelable";
  tableId: string;
  ranges: string[];
  rawContent: string;
}

export interface UnmodelableTabRuleRecord {
  kind: "unmodelable";
  sheetGid: number;
  ranges: string[];
  rawContent: string;
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

// A handle is readable content, not a hash: packages/utils has no hash function.
export const RuleRecords = {
  build(tabs: RuleRecordsTab[]): RuleRecordSet {
    return tabs.reduce<RuleRecordSet>(
      (recordSet, tab) => {
        tab.rules.forEach((rule) => {
          const table = tab.tables.find((candidate) =>
            rule.ranges.some((range) => touchesTable(range, candidate)),
          );
          if (table === undefined) {
            const group = (recordSet.tabs[tab.tabName] ??= {});
            addRecord(group, tabRecord(rule, tab.sheetGid));
            return;
          }
          const group = (recordSet.tables[table.tableKey] ??= {});
          addRecord(group, tableRecord(rule, table));
        });
        return recordSet;
      },
      { tables: {}, tabs: {} },
    );
  },
};

function touchesTable(range: GridRangeProps, table: RuleRecordsTable): boolean {
  const { startCol, endCol } = tableColumnBounds(table);
  const overlapsRows =
    range.startRowIndex < table.endRowIndex &&
    (range.endRowIndex ?? Infinity) > table.layout.headerTopLeft.rowIndex;
  const overlapsColumns =
    (range.startColumnIndex ?? 0) < endCol &&
    (range.endColumnIndex ?? Infinity) > startCol;
  return overlapsRows && overlapsColumns;
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

function addRecord<RR>(
  group: Record<string, RR>,
  { handle, record }: HandledRecord<RR>,
): void {
  let suffixed = handle;
  let repeat = 1;
  while (group[suffixed] !== undefined) {
    repeat += 1;
    suffixed = `${handle}-${repeat}`;
  }
  group[suffixed] = record;
}

function tabRecord(
  rule: ConditionalFormatRule,
  sheetGid: number,
): HandledRecord<TabRuleRecord | UnmodelableTabRuleRecord> {
  const ranges = rule.ranges.map(rangeA1);
  if (rule.kind === "unmodelable") {
    const { rawContent } = rule;
    return {
      handle: unmodelableHandle(ranges, rawContent),
      record: { kind: "unmodelable", sheetGid, ranges, rawContent },
    };
  }
  const format = recordFormat(rule.format);
  return {
    handle: handleOf(ranges.join(","), rule.condition, format),
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
  const ranges = rule.ranges.map(rangeA1);
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
  const condition = braceCondition(rule.condition, table, columnNames);
  if (condition === undefined) return undefined;
  let columns: string[] | "all" = columnNames;
  if (columnNames.length === table.layout.columnNamesLeftToRight.length) {
    columns = "all";
  }
  const format = recordFormat(rule.format);
  return {
    handle: handleOf(columnsText(columns), condition, format),
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
  return table.layout.columnNamesLeftToRight.filter((_, offset) =>
    offsets.has(offset),
  );
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
  if (endRowIndex === undefined || endRowIndex > table.endRowIndex) {
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

function braceCondition(
  condition: BooleanCondition,
  table: RuleRecordsTable,
  columnNames: string[],
): BooleanCondition | undefined {
  if (condition.type !== "CUSTOM_FORMULA") return condition;
  const formula = BraceFormula.decompileForRule({
    tableKey: table.tableKey,
    layout: table.layout,
    formula: condition.formula,
    isOneColumn: columnNames.length === 1,
  });
  if (formula === undefined) return undefined;
  return { type: "CUSTOM_FORMULA", formula };
}

function columnsText(columns: string[] | "all"): string {
  if (columns === "all") return "all";
  return columns.join("+");
}

function handleOf(
  scope: string,
  condition: BooleanCondition,
  format: RuleRecordFormat,
): string {
  return `${scope} | ${conditionText(condition)} | ${formatText(format)}`;
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

// An absent bound is open, so it is left off the A1 text.
function rangeA1({
  startRowIndex,
  endRowIndex,
  startColumnIndex,
  endColumnIndex,
}: GridRangeProps): string {
  const startCol = schema.columnLetter(startColumnIndex ?? SheetIndex.col(0));
  let endCol = "";
  if (endColumnIndex !== undefined) {
    endCol = schema.columnLetter(SheetIndex.col(endColumnIndex - 1));
  }
  let endRow = "";
  if (endRowIndex !== undefined) endRow = String(endRowIndex);
  return `${startCol}${startRowIndex + 1}:${endCol}${endRow}`;
}
