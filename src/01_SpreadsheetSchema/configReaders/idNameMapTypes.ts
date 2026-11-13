export interface IdNameMapTable {
  name: string;
  sheetGid: number;
  columns?: Record<string, string>;
}
export interface IdNameMapTab {
  name: string;
  title: string;
  tableIds: string[];
}
// Tables are keyed by Table ID and tabs by GID, the two keys a live spreadsheet keeps unique.
export interface IdNameMapGeneric {
  tables: Record<string, IdNameMapTable>;
  tabs: Record<string, IdNameMapTab>;
}
export interface ColumnReference {
  tableId: string;
  columnId: string;
}

export interface MapTablesByLiveness {
  live: IdNameMapTable[];
  gone: IdNameMapTable[];
}

// The map's Tables split by whether the sheet still holds them.
export function mapTablesByLiveness(
  { tables }: IdNameMapGeneric,
  liveTableIds: ReadonlySet<string>,
): MapTablesByLiveness {
  return Object.entries(tables).reduce<MapTablesByLiveness>(
    (split, [tableId, mapTable]) => {
      if (liveTableIds.has(tableId)) split.live.push(mapTable);
      else split.gone.push(mapTable);
      return split;
    },
    { live: [], gone: [] },
  );
}

// Each Table's ID by the codebase name alterations.ts keys it by.
export function tableIdsByName({
  tables,
}: IdNameMapGeneric): Map<string, string> {
  return new Map(
    Object.entries(tables).map(([tableId, { name }]) => [name, tableId]),
  );
}
