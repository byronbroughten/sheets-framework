import type { ColumnTypeBaseline } from "../../columnBaseline";

// `PendingAlterations` keyed by whatever the map holds, widened since the planner reads every entry alike.
export interface WideRule {
  columns: string[] | "all";
  condition: { values?: string[] };
}

export interface WideColumnEntry {
  deleted?: true;
  before?: string;
  formula?: string;
  header?: string;
  columnType?: ColumnTypeBaseline;
  emptyValueAllowed?: boolean;
}

export interface WideTableEntry {
  deleted?: true;
  name?: string;
  letApiAccess?: boolean;
  columns?: Record<string, WideColumnEntry | undefined>;
  newColumns?: WideColumnEntry[];
  conditionalFormats?: {
    add?: WideRule[];
    replace?: Record<string, WideRule | undefined>;
  };
}

export interface WideNewTable {
  name: string;
  columns: { formula?: string }[];
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
  newTabs?: (WideTabEntry & { tables: WideNewTable[] })[];
}

export function entriesOf<VL>(
  record: Record<string, VL | undefined> | undefined,
): [string, VL][] {
  return Object.entries(record ?? {}).flatMap(([key, value]) => {
    if (value === undefined) return [];
    return [[key, value] as [string, VL]];
  });
}
