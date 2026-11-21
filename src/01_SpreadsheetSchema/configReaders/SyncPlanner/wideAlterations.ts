import type { ColumnTypeBaseline } from "../../columnBaseline";

// `PendingAlterations` keyed by whatever the map holds, widened since the planner reads every entry alike.
export interface WideRule {
  columns: string[] | "all";
  condition: { type: string; values?: string[] };
  format: WideRuleFormat;
}

export interface WideRuleFormat {
  backgroundColor?: string;
  foregroundColor?: string;
  bold?: boolean;
  italic?: boolean;
  strikethrough?: boolean;
}

export interface WideColumnEntry {
  deleted?: true;
  formula?: string;
  header?: string;
  columnType?: ColumnTypeBaseline;
  emptyValueAllowed?: boolean;
}

export interface WideNewColumn extends WideColumnEntry {
  before?: string;
}

export interface WideNewColumnGroup {
  columns: WideNewTableColumn[];
  before?: string;
}

export interface WideColumnGroupMove {
  columns: string[];
  before: string;
}

export interface WideTableEntry {
  deleted?: true;
  name?: string;
  letApiAccess?: boolean;
  columns?: Record<string, WideColumnEntry | undefined>;
  newColumns?: WideNewColumnGroup[];
  moves?: WideColumnGroupMove[];
  conditionalFormats?: {
    add?: WideRule[];
    remove?: string[];
    replace?: Record<string, WideRule | undefined>;
  };
}

export interface WideNewTableColumn {
  header: string;
  formula?: string;
  columnType?: ColumnTypeBaseline;
  emptyValueAllowed?: boolean;
}

export interface WideNewTable {
  name: string;
  columns: WideNewTableColumn[];
  conditionalFormats?: { add?: WideRule[] };
  gap?: 0 | 1;
}

export interface WideTabEntry {
  deleted?: true;
  before?: string;
  title?: string;
  color?: string;
  frozenRows?: number;
  frozenColumns?: number;
  hidden?: boolean;
  newTables?: WideNewTable[];
}

export interface WideAlterations {
  tables?: Record<string, WideTableEntry | undefined>;
  tabs?: Record<string, WideTabEntry | undefined>;
  newTabs?: (WideTabEntry & { title: string; tables: WideNewTable[] })[];
}

// A `moves` or `newColumns` group's column by its group's index and its own within the group.
export interface GroupColumnAt {
  index: number;
  columnIndex: number;
}

// A rule a Table entry adds or writes over a handle, by where it sits in the entry.
export interface WideRuleAt {
  path: string;
  rule: WideRule;
}

// A new column with the `before` its `newColumns` group gives it, by where it sits in its Table entry.
export interface WideNewColumnAt {
  newColumn: WideNewColumn;
  // As `newColumns[0].columns[1]`.
  path: string;
  // Its group's `before`'s path, as `newColumns[0].before`.
  beforePath: string;
}

// A `newColumns` group's `before`, by its path, as `newColumns[0].before`.
export interface WideNewColumnsBefore {
  before: string;
  path: string;
}

// A `newColumns` group's path in its Table entry, as `newColumns[0]`.
export function newColumnPath(index: number): string {
  return `newColumns[${index}]`;
}

// A `newColumns` group's `before`'s path in its Table entry, as `newColumns[0].before`.
export function newColumnBeforePath(index: number): string {
  return `${newColumnPath(index)}.before`;
}

// A `newColumns` group's column's path in its Table entry, as `newColumns[0].columns[1]`.
export function newColumnGroupColumnPath({
  index,
  columnIndex,
}: GroupColumnAt): string {
  return `${newColumnPath(index)}.columns[${columnIndex}]`;
}

// Each column a Table entry's `newColumns` create, in order, each taking its group's `before`.
export function newColumnsOf({
  newColumns,
}: WideTableEntry): WideNewColumnAt[] {
  return (newColumns ?? []).flatMap(({ columns, before }, index) => {
    const beforePath = newColumnBeforePath(index);
    return columns.map((column, columnIndex) => ({
      newColumn: { ...column, ...(before === undefined ? {} : { before }) },
      path: newColumnGroupColumnPath({ index, columnIndex }),
      beforePath,
    }));
  });
}

// One per `newColumns` group that has a `before`.
export function newColumnsBeforesOf({
  newColumns,
}: WideTableEntry): WideNewColumnsBefore[] {
  return (newColumns ?? []).flatMap(({ before }, index) => {
    if (before === undefined) return [];
    return [{ before, path: newColumnBeforePath(index) }];
  });
}

// A `moves` group's path in its Table entry, as `moves[0]`.
export function movePath(index: number): string {
  return `moves[${index}]`;
}

// A `moves` group's column's path in its Table entry, as `moves[0].columns[1]`.
export function moveColumnPath({ index, columnIndex }: GroupColumnAt): string {
  return `${movePath(index)}.columns[${columnIndex}]`;
}

// A `moves` group's target's path in its Table entry, as `moves[0].before`.
export function moveBeforePath(index: number): string {
  return `${movePath(index)}.before`;
}

export function isMovePath(path: string): boolean {
  return /^moves\[\d+\]$/.test(path);
}

// A `conditionalFormats.replace` entry's path in its Table entry, as `conditionalFormats.replace["amount | =TRUE"]`.
export function ruleReplacementPath(handle: string): string {
  return `conditionalFormats.replace["${handle}"]`;
}

export function rulesOf({ conditionalFormats }: WideTableEntry): WideRuleAt[] {
  return [
    ...(conditionalFormats?.add ?? []).map((rule, index) => ({
      path: `conditionalFormats.add[${index}]`,
      rule,
    })),
    ...entriesOf(conditionalFormats?.replace).map(([handle, rule]) => ({
      path: ruleReplacementPath(handle),
      rule,
    })),
  ];
}

export function entriesOf<VL>(
  record: Record<string, VL | undefined> | undefined,
): [string, VL][] {
  return Object.entries(record ?? {}).flatMap(([key, value]) => {
    if (value === undefined) return [];
    return [[key, value] as [string, VL]];
  });
}
