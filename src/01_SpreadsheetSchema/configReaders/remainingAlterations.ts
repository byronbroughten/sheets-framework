import {
  entriesOf,
  type WideAlterations,
  type WideColumnEntry,
  type WideTableEntry,
} from "./SyncPlanner/wideAlterations";
import type { SyncSubject } from "./syncRunResult";

// The alterations still pending once `landed` went out; the refuser lets only a Table's, a column's or a tab's trait through, so any other subject is a bug.
export function remainingAlterations(
  alterations: WideAlterations,
  landed: SyncSubject[],
): WideAlterations {
  const landedKeys = new Set(landed.map(landedTraitKey));
  return Object.fromEntries(
    Object.entries(alterations).flatMap(([key, value]) => {
      if (key === "tables") {
        return keepIfNotEmpty(
          key,
          remainingTables(alterations.tables, landedKeys),
        );
      }
      if (key === "tabs") {
        return keepIfNotEmpty(key, remainingTabs(alterations.tabs, landedKeys));
      }
      return [[key, value]];
    }),
  );
}

function landedTraitKey(subject: SyncSubject): string {
  if ("tab" in subject && subject.path) {
    return traitKey(tabKeyPrefix, subject.tab, subject.path);
  }
  if (!("table" in subject) || !subject.path) {
    throw new Error(
      `A landed subject must be a Table, column or tab trait: ${JSON.stringify(subject)}`,
    );
  }
  if (!subject.column) return traitKey(subject.table, subject.path);
  return traitKey(subject.table, subject.column, subject.path);
}

// Marks a tab's key, so a tab and a Table sharing a name can't collide.
const tabKeyPrefix = "tab:";

// A Table's trait is [table, trait]; a column's is [table, column, trait]; a tab's is ["tab:", tab, trait].
function traitKey(...parts: string[]): string {
  return JSON.stringify(parts);
}

function keepIfNotEmpty(key: string, record: object): [string, object][] {
  return Object.keys(record).length === 0 ? [] : [[key, record]];
}

function remainingTabs(
  tabs: WideAlterations["tabs"],
  landedKeys: Set<string>,
): Record<string, object> {
  return remainingEntries(tabs, landedKeys, tabKeyPrefix);
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
      if (key !== "columns") {
        return landedKeys.has(traitKey(tableName, key)) ? [] : [[key, value]];
      }
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
  return remainingEntries(columns, landedKeys, tableName);
}

// A tab's or a column's entries, each keeping only its traits not landed, keyed under `prefix`.
function remainingEntries(
  entries: Record<string, object | undefined> | undefined,
  landedKeys: Set<string>,
  prefix: string,
): Record<string, object> {
  return Object.fromEntries(
    entriesOf(entries).flatMap(([entryName, entry]) =>
      keepIfNotEmpty(
        entryName,
        Object.fromEntries(
          Object.entries(entry).filter(
            ([trait]) => !landedKeys.has(traitKey(prefix, entryName, trait)),
          ),
        ),
      ),
    ),
  );
}
