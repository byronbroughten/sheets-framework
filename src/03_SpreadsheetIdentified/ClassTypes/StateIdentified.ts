import type { RgbColor } from "../../00_Source/RawSource/RgbColor";
import type { Value, ValueName } from "../../01_SpreadsheetSchema/valueSchemas";

// Mirrors the raw queued entry's optional pair, so a run can write either or both.
export interface CellChange<VN extends ValueName = ValueName> {
  value?: Value<VN>;
  backgroundColor?: RgbColor;
}

export interface StateIdentified {
  tables: TablesStateIdentified;
  // What `ss.sheetMeta(gid).primary` preps before its Table is known.
  tableBeforePropertiesBySheet: Map<SheetId, TableStateIdentified>;
}

export type TablesStateIdentified = Map<TableId, TableStateIdentified>;

export interface TableStateIdentified {
  fetchQueue: TableFetchQueueIdentified;
}

export interface TableFetchQueueIdentified {
  targets: FetchTargetIdentified[];
}

export function emptyTableStateIdentified(): TableStateIdentified {
  return { fetchQueue: { targets: [] } };
}

type SheetId = number;
type TableId = string;

interface FullRowTarget {
  kind: "fullRow";
  row: number;
}
interface FullColumnTarget {
  kind: "fullDataColumn";
  column: string;
}
interface CellTarget {
  kind: "singleCell";
  row: number;
  column: string;
}

export type FetchTargetIdentified =
  FullRowTarget | FullColumnTarget | CellTarget;
