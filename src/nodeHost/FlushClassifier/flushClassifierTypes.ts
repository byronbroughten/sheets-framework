import type {
  GoogleRequest,
  SheetsHttpRead,
} from "../../00_Source/GoogleSheets/GoogleSheetsAPI.ts";
import type { ForcedTrait } from "../../01_SpreadsheetSchema/configReaders/syncRunResult.ts";

export type GoogleGridRange = GoogleAppsScript.Sheets.Schema.GridRange;
export type GoogleGridCoordinate =
  GoogleAppsScript.Sheets.Schema.GridCoordinate;
export type GoogleCellData = GoogleAppsScript.Sheets.Schema.CellData;
export type ExtendedValue = GoogleAppsScript.Sheets.Schema.ExtendedValue;
export type RequestVerb = keyof GoogleRequest;

export type FlushRead = SheetsHttpRead;

export interface GridCell {
  sheetGid: number;
  rowIndex: number;
  colIndex: number;
}

export interface GridSize {
  rowCount: number;
  columnCount: number;
}

export type ShiftDimension = "ROWS" | "COLUMNS";

export interface IndexSpan {
  startIndex: number;
  endIndex: number;
}

export interface BoundedRange {
  sheetGid: number;
  rows: IndexSpan;
  cols: IndexSpan;
}

export type CellOrigin =
  | { kind: "read"; cell: GridCell }
  // Inserted, appended, on an added tab, or emptied by a cut.
  | { kind: "created" }
  | { kind: "written" }
  // Sorted earlier in the flush, so which read cell it holds is unknown.
  | { kind: "reordered" };

export type CellBack = GridCell | Exclude<CellOrigin, { kind: "read" }>;

// One request's effect on the grid, as the cells it left mapped back and the ranges it found mapped on.
export interface TrackerLayer {
  back(cell: GridCell): CellBack;
  // Undefined once the request split the range, so it no longer stands as one range.
  forward(range: BoundedRange): BoundedRange | undefined;
}

export interface DimensionShift {
  sheetGid: number;
  dimension: ShiftDimension;
  startIndex: number;
  // Positive for an insert, negative for a delete.
  count: number;
  // The span across the other dimension a range shift stays inside; undefined for whole rows or columns.
  band: IndexSpan | undefined;
}

export interface DimensionMove extends IndexSpan {
  sheetGid: number;
  dimension: ShiftDimension;
  // Counted before the moved span is lifted out, as Google counts it.
  destinationIndex: number;
}

export interface RangeCut {
  source: BoundedRange;
  destination: GridCell;
}

export type HeldRow =
  { rowId: string } | { sheetRow: number; nameValue: string | undefined };

export interface HeldCell {
  sheetGid: number;
  // Undefined where no read covered the sheet's Tables.
  table: string | undefined;
  column: string;
  row: HeldRow;
}

// Where a held cell's column stood when the run last read its Table: what a send snapshot reads.
export interface HeldColumn {
  sheetGid: number;
  tabTitle: string;
  table: string;
  tableId: string;
  header: string;
  colIndex: number;
  columnIdRowIndex: number;
  // Undefined where the Table has no ID column.
  idColIndex: number | undefined;
  tabColumnCount: number;
}

export type CellChange =
  | { kind: "overwrite"; before: string; after: string }
  | { kind: "delete"; before: string };

export type HeldItem =
  | {
      kind: "enteredData";
      requestIndex: number;
      verb: RequestVerb;
      cell: HeldCell;
      change: CellChange;
    }
  | {
      kind: "unreadCell";
      requestIndex: number;
      verb: RequestVerb;
      cell: HeldCell;
    }
  | {
      kind: "unmeasured";
      requestIndex: number;
      verb: RequestVerb;
      detail: string;
    }
  | {
      kind: "findReplace";
      requestIndex: number;
      find: string;
      replacement: string;
    }
  | { kind: "dropdownReset"; requestIndex: number; table: string }
  | { kind: "forcedTrait"; trait: ForcedTrait }
  | { kind: "unclassified"; requestIndex: number; verb: RequestVerb };

// The held items that name a cell; every other kind is undone only by version history.
export type HeldCellItem = Extract<
  HeldItem,
  { kind: "enteredData" | "unreadCell" }
>;

// What a judged cell holds after the request: a value, nothing, or no cell at all.
export type CellAfter =
  | { kind: "value"; value: ExtendedValue }
  | { kind: "cleared" }
  | { kind: "deleted" }
  | { kind: "unknown"; label: string };
