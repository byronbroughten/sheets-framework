import {
  entriesOf,
  type WideAlterations,
  type WideColumnEntry,
  type WideTableEntry,
} from "./SyncPlanner/wideAlterations";
import type { SyncSubject } from "./syncRunResult";

// The alterations still pending once `landed` went out; the refuser lets only a Table's or a column's trait through, so any other subject is a bug.
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
      `A landed subject must be a Table or column trait: ${JSON.stringify(subject)}`,
    );
  }
  if (!subject.column) return tableTraitKey(subject.table, subject.path);
  return traitKey(subject.table, subject.column, subject.path);
}

function tableTraitKey(table: string, trait: string): string {
  return JSON.stringify([table, trait]);
}

function traitKey(table: string, column: string, trait: string): string {
  return JSON.stringify([table, column, trait]);
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
      if (key !== "columns") {
        return landedKeys.has(tableTraitKey(tableName, key))
          ? []
          : [[key, value]];
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
