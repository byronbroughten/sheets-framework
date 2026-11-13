import {
  entriesOf,
  type WideAlterations,
  type WideColumnEntry,
  type WideTableEntry,
} from "./SyncPlanner/wideAlterations";
import type { SyncSubject } from "./syncRunResult";

// The alterations still pending once `landed` went out; the refuser lets only a Table's or column's trait through, so a tab subject is a bug.
export function remainingAlterations(
  alterations: WideAlterations,
  landed: SyncSubject[],
): WideAlterations {
  const landedKeys = new Set(landed.map(landedTraitKey));
  return Object.fromEntries(
    Object.entries(alterations).flatMap(([key, value]) => {
      if (key !== "tables") return [[key, value]];
      return keepIfNotEmpty(
        key,
        remainingTables(alterations.tables, landedKeys),
      );
    }),
  );
}

function landedTraitKey(subject: SyncSubject): string {
  if (!("table" in subject) || !subject.path) {
    throw new Error(
      `A landed subject must be a Table's or column's trait: ${JSON.stringify(subject)}`,
    );
  }
  if (!subject.column) return traitKey(subject.table, subject.path);
  return traitKey(subject.table, subject.column, subject.path);
}

// A Table's trait is [table, trait]; a column's is [table, column, trait].
function traitKey(...parts: string[]): string {
  return JSON.stringify(parts);
}

function keepIfNotEmpty(key: string, record: object): [string, object][] {
  return Object.keys(record).length === 0 ? [] : [[key, record]];
}

function remainingTables(
  tables: WideAlterations["tables"],
  landedKeys: Set<string>,
): Record<string, WideTableEntry> {
  return Object.fromEntries(
    entriesOf(tables).flatMap(([tableName, entry]) =>
      keepIfNotEmpty(
        tableName,
        remainingTableEntry(tableName, entry, landedKeys),
      ),
    ),
  );
}

function remainingTableEntry(
  tableName: string,
  entry: WideTableEntry,
  landedKeys: Set<string>,
): WideTableEntry {
  return Object.fromEntries(
    Object.entries(entry).flatMap(([key, value]) => {
      if (landedKeys.has(traitKey(tableName, key))) return [];
      if (key !== "columns") return [[key, value]];
      return keepIfNotEmpty(
        key,
        remainingColumns(tableName, entry.columns, landedKeys),
      );
    }),
  );
}

function remainingColumns(
  tableName: string,
  columns: WideTableEntry["columns"],
  landedKeys: Set<string>,
): Record<string, WideColumnEntry> {
  return Object.fromEntries(
    entriesOf(columns).flatMap(([columnName, entry]) =>
      keepIfNotEmpty(
        columnName,
        remainingColumnEntry(tableName, columnName, entry, landedKeys),
      ),
    ),
  );
}

function remainingColumnEntry(
  tableName: string,
  columnName: string,
  entry: WideColumnEntry,
  landedKeys: Set<string>,
): WideColumnEntry {
  return Object.fromEntries(
    Object.entries(entry).filter(
      ([trait]) => !landedKeys.has(traitKey(tableName, columnName, trait)),
    ),
  );
}
