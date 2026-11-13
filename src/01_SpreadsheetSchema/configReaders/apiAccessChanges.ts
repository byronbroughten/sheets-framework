import type { IdNameMapGeneric } from "./idNameMapTypes";
import type { PlanProps, TableReading, TraitReadings } from "./SyncPlanner";
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
  // The previous one: a Table has API access when it lists columns.
  idNameMap: IdNameMapGeneric;
}

export function apiAccessChanges({
  alterations,
  idNameMap,
}: ApiAccessChangesProps): ApiAccessChanges {
  const tableIdsByName = new Map(
    Object.entries(idNameMap.tables).map(([tableId, { name }]) => [
      name,
      tableId,
    ]),
  );
  return entriesOf(alterations.tables).reduce<ApiAccessChanges>(
    (changes, [table, { letApiAccess }]) => {
      const tableId = tableIdsByName.get(table);
      if (tableId === undefined || letApiAccess === undefined) return changes;
      const hasApiAccess = idNameMap.tables[tableId]?.columns !== undefined;
      if (letApiAccess && !hasApiAccess) {
        changes.grants.push({ tableId, table });
      }
      if (!letApiAccess && hasApiAccess) changes.revokedTableIds.push(tableId);
      return changes;
    },
    { grants: [], revokedTableIds: [] },
  );
}

// API access isn't a sheet trait, so live carries the previous map and a grant or revoke never stops.
export function withApiAccessReadings(
  readings: PlanProps,
  idNameMap: IdNameMapGeneric,
): PlanProps {
  return {
    ...readings,
    previous: withApiAccess(readings.previous, idNameMap),
    live: withApiAccess(readings.live, idNameMap),
  };
}

function withApiAccess(
  readings: TraitReadings,
  idNameMap: IdNameMapGeneric,
): TraitReadings {
  const tables = Object.values(idNameMap.tables).reduce<
    TraitReadings["tables"]
  >(
    (byName, { name, columns }) => {
      const reading: TableReading = byName[name] ?? { traits: {}, columns: {} };
      byName[name] = {
        ...reading,
        traits: { ...reading.traits, letApiAccess: columns !== undefined },
      };
      return byName;
    },
    { ...readings.tables },
  );
  return { ...readings, tables };
}
