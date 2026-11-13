import {
  type IdNameMapGeneric,
  type IdNameMapTable,
  mapTablesByLiveness,
  tableIdsByName,
} from "./idNameMapTypes";
import type { PlanReadings, TableReading, TraitReadings } from "./SyncPlanner";
import { entriesOf, type WideAlterations } from "./SyncPlanner/wideAlterations";

export interface ApiAccessGrant {
  tableId: string;
  // The codebase name alterations.ts keys it by.
  table: string;
}

// What this run's letApiAccess entries would change, by Table ID; an entry agreeing with the map changes nothing.
export interface ApiAccessChanges {
  grants: ApiAccessGrant[];
  revokedTableIds: string[];
}

interface ApiAccessChangesProps {
  alterations: WideAlterations;
  // The previous one, since API access is read from it.
  idNameMap: IdNameMapGeneric;
}

// A map Table has API access when it lists columns.
export function hasApiAccess(mapTable: IdNameMapTable): boolean {
  return mapTable.columns !== undefined;
}

export function apiAccessChanges({
  alterations,
  idNameMap,
}: ApiAccessChangesProps): ApiAccessChanges {
  const tableIds = tableIdsByName(idNameMap);
  return entriesOf(alterations.tables).reduce<ApiAccessChanges>(
    (changes, [table, { letApiAccess }]) => {
      const tableId = tableIds.get(table);
      if (tableId === undefined || letApiAccess === undefined) return changes;
      const mapTable = idNameMap.tables[tableId];
      const isGranted = mapTable !== undefined && hasApiAccess(mapTable);
      if (letApiAccess && !isGranted) changes.grants.push({ tableId, table });
      if (!letApiAccess && isGranted) changes.revokedTableIds.push(tableId);
      return changes;
    },
    { grants: [], revokedTableIds: [] },
  );
}

interface ApiAccessReadingsProps {
  readings: PlanReadings;
  previousIdNameMap: IdNameMapGeneric;
  liveTableIds: ReadonlySet<string>;
}

// API access isn't a sheet trait, so live carries the previous map and a grant or revoke never stops; a Table a person deleted stays out of live.
export function withApiAccessReadings({
  readings,
  previousIdNameMap,
  liveTableIds,
}: ApiAccessReadingsProps): PlanReadings {
  const { live } = mapTablesByLiveness(previousIdNameMap, liveTableIds);
  return {
    previous: withApiAccess(
      readings.previous,
      Object.values(previousIdNameMap.tables),
    ),
    live: withApiAccess(readings.live, live),
  };
}

function withApiAccess(
  readings: TraitReadings,
  mapTables: IdNameMapTable[],
): TraitReadings {
  const tables = mapTables.reduce<TraitReadings["tables"]>(
    (byName, mapTable) => {
      const reading: TableReading = byName[mapTable.name] ?? {
        traits: {},
        columns: {},
      };
      byName[mapTable.name] = {
        ...reading,
        traits: { ...reading.traits, letApiAccess: hasApiAccess(mapTable) },
      };
      return byName;
    },
    { ...readings.tables },
  );
  return { ...readings, tables };
}
