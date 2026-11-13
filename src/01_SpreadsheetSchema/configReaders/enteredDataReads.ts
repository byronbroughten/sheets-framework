import type { IdNameMapGeneric } from "./idNameMapTypes";
import type { IntendedTrait } from "./SyncPlanner";
import { entriesOf, type WideAlterations } from "./SyncPlanner/wideAlterations";

export interface EnteredDataColumn {
  tableId: string;
  columnId: string;
}

// What sync reads before its flush, so the entered-data gate can judge each cell a deletion or a formula would overwrite.
export interface EnteredDataReads {
  tabGids: number[];
  tableIds: string[];
  columns: EnteredDataColumn[];
}

interface EnteredDataReadsProps {
  alterations: WideAlterations;
  sentTraits: IntendedTrait[];
  // The previous one, since the alterations name what it records.
  idNameMap: IdNameMapGeneric;
}

// A name the map can't resolve is left out: the entry refusals have already refused any that matter.
export function enteredDataReads({
  alterations,
  sentTraits,
  idNameMap,
}: EnteredDataReadsProps): EnteredDataReads {
  const tableIdsByName = new Map(
    Object.entries(idNameMap.tables).map(([tableId, { name }]) => [
      name,
      tableId,
    ]),
  );
  function columnOf(
    table: string,
    column: string,
  ): EnteredDataColumn | undefined {
    const tableId = tableIdsByName.get(table);
    if (tableId === undefined) return undefined;
    const columnId = idNameMap.tables[tableId]?.columns?.[column];
    if (columnId === undefined) return undefined;
    return { tableId, columnId };
  }
  const tableEntries = entriesOf(alterations.tables);
  const deletedColumns = tableEntries.flatMap(([table, entry]) =>
    entriesOf(entry.columns).flatMap(([column, columnEntry]) => {
      if (columnEntry.deleted !== true) return [];
      const read = columnOf(table, column);
      return read === undefined ? [] : [read];
    }),
  );
  const formulaColumns = sentTraits.flatMap((intended) => {
    if (intended.kind !== "column" || intended.trait !== "formula") return [];
    const read = columnOf(intended.table, intended.column);
    return read === undefined ? [] : [read];
  });
  const tableIds = tableEntries.flatMap(([table, entry]) => {
    if (entry.deleted !== true) return [];
    const tableId = tableIdsByName.get(table);
    return tableId === undefined ? [] : [tableId];
  });
  const tabGidsByName = new Map(
    Object.entries(idNameMap.tabs).map(([gid, { name }]) => [
      name,
      Number(gid),
    ]),
  );
  const tabGids = entriesOf(alterations.tabs).flatMap(([tab, entry]) => {
    if (entry.deleted !== true) return [];
    const gid = tabGidsByName.get(tab);
    return gid === undefined ? [] : [gid];
  });
  return {
    tabGids,
    tableIds,
    columns: [...deletedColumns, ...formulaColumns],
  };
}
