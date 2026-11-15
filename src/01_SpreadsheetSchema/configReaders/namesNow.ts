import type { IdNameMapGeneric, IdNameMapTable } from "./idNameMapTypes";
import type { SyncMovedName } from "./syncRunResult";

interface ColumnNow {
  columnId: string;
  name: string;
}

export interface TableNow {
  tableId: string;
  name: string;
  // Keyed by each column's previous name.
  columns: Map<string, ColumnNow>;
}

interface TabNow {
  sheetGid: string;
  name: string;
  title: string;
}

// Each Table and tab both maps hold, keyed by its previous name, with the name the regen gives it now.
export interface NamesNow {
  tables: Map<string, TableNow>;
  tabs: Map<string, TabNow>;
}

interface IdNameMaps {
  previous: IdNameMapGeneric;
  regenerated: IdNameMapGeneric;
}

interface TableInBoth {
  tableId: string;
  previous: IdNameMapTable;
  regenerated: IdNameMapTable;
}

export function namesNow(maps: IdNameMaps): NamesNow {
  const tables = new Map<string, TableNow>();
  for (const { tableId, previous, regenerated } of tablesInBoth(maps)) {
    const namesById = new Map(
      Object.entries(regenerated.columns ?? {}).map(([name, columnId]) => [
        columnId,
        name,
      ]),
    );
    const columns = new Map<string, ColumnNow>();
    for (const [previousName, columnId] of Object.entries(
      previous.columns ?? {},
    )) {
      const name = namesById.get(columnId);
      if (name !== undefined) columns.set(previousName, { columnId, name });
    }
    tables.set(previous.name, { tableId, name: regenerated.name, columns });
  }
  const tabs = new Map<string, TabNow>();
  for (const [sheetGid, previousTab] of Object.entries(maps.previous.tabs)) {
    const tab = maps.regenerated.tabs[sheetGid];
    if (tab === undefined) continue;
    tabs.set(previousTab.name, { sheetGid, name: tab.name, title: tab.title });
  }
  return { tables, tabs };
}

// Every column name both maps hold for one Table that now names another column ID.
export function movedNames(maps: IdNameMaps): SyncMovedName[] {
  return tablesInBoth(maps).flatMap(({ previous, regenerated }) =>
    Object.entries(previous.columns ?? {}).flatMap(
      ([column, wasId]): SyncMovedName[] => {
        const nowId = regenerated.columns?.[column];
        if (nowId === undefined || nowId === wasId) return [];
        return [{ table: regenerated.name, column, wasId, nowId }];
      },
    ),
  );
}

function tablesInBoth({ previous, regenerated }: IdNameMaps): TableInBoth[] {
  return Object.entries(previous.tables).flatMap(
    ([tableId, previousTable]): TableInBoth[] => {
      const table = regenerated.tables[tableId];
      if (table === undefined) return [];
      return [{ tableId, previous: previousTable, regenerated: table }];
    },
  );
}
