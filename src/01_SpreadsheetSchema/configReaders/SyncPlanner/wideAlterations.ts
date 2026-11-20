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
  before?: string;
  formula?: string;
  header?: string;
  columnType?: ColumnTypeBaseline;
  emptyValueAllowed?: boolean;
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
  newColumns?: WideColumnEntry[];
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

// A `moves` group's column by its group's index and its own within the group.
export interface MoveColumnAt {
  index: number;
  columnIndex: number;
}

// A rule a Table entry adds or writes over a handle, by where it sits in the entry.
export interface WideRuleAt {
  path: string;
  rule: WideRule;
}

// A `newColumns` entry's path in its Table entry, as `newColumns[0]`.
export function newColumnPath(index: number): string {
  return `newColumns[${index}]`;
}

// A `moves` group's path in its Table entry, as `moves[0]`.
export function movePath(index: number): string {
  return `moves[${index}]`;
}

// A `moves` group's column's path in its Table entry, as `moves[0].columns[1]`.
export function moveColumnPath({ index, columnIndex }: MoveColumnAt): string {
  return `${movePath(index)}.columns[${columnIndex}]`;
}

// A `moves` group's target's path in its Table entry, as `moves[0].before`.
export function moveBeforePath(index: number): string {
  return `${movePath(index)}.before`;
}

export function isMovePath(path: string): boolean {
  return /^moves\[\d+\]$/.test(path);
}

export function rulesOf({ conditionalFormats }: WideTableEntry): WideRuleAt[] {
  return [
    ...(conditionalFormats?.add ?? []).map((rule, index) => ({
      path: `conditionalFormats.add[${index}]`,
      rule,
    })),
    ...entriesOf(conditionalFormats?.replace).map(([handle, rule]) => ({
      path: `conditionalFormats.replace["${handle}"]`,
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
