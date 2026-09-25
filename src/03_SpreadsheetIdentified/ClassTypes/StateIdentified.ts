import type { RgbColor } from "../../00_Source/RawSource/RgbColor";
import type { Value, ValueName } from "../../01_SpreadsheetSchema/valueSchemas";

// Mirrors the raw queued entry's optional pair, so a run can write either or both.
export interface CellChange<VN extends ValueName = ValueName> {
  value?: Value<VN>;
  backgroundColor?: RgbColor;
}

export interface StateIdentified {
  sheets: SheetsStateIdentified;
}

export type SheetsStateIdentified = Map<SheetId, SheetStateIdentified>;

export interface SheetStateIdentified {
  fetchQueue: SheetFetchQueueIdentified;
}

export interface SheetFetchQueueIdentified {
  targets: FetchTargetIdentified[];
  gatherConditionalFormats: boolean;
  gatherEditProtections: boolean;
}

export function emptySheetFetchQueueIdentified(): SheetFetchQueueIdentified {
  return {
    targets: [],
    gatherConditionalFormats: false,
    gatherEditProtections: false,
  };
}

type SheetId = number;

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
