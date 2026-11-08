import type { GoogleRequest } from "../../00_Source/GoogleSheets/GoogleSheetsAPI";
import type { ForcedTrait } from "../../05_Operators/syncRunResult";

export type GoogleGridRange = GoogleAppsScript.Sheets.Schema.GridRange;
export type GoogleGridCoordinate =
  GoogleAppsScript.Sheets.Schema.GridCoordinate;
export type GoogleCellData = GoogleAppsScript.Sheets.Schema.CellData;
export type ExtendedValue = GoogleAppsScript.Sheets.Schema.ExtendedValue;
export type RequestVerb = keyof GoogleRequest;

export interface FlushRead {
  // Google trims trailing empty rows and cells from a reply, so what a read covers comes from the ranges it asked for.
  gridRanges: GoogleGridRange[];
  spreadsheet: GoogleAppsScript.Sheets.Schema.Spreadsheet;
}

export type DropdownStyleReset = "ask" | "allow";

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

export type HeldRow =
  { rowId: string } | { sheetRow: number; nameValue: string | undefined };

export interface HeldCell {
  sheetGid: number;
  // Undefined where no read covered the sheet's Tables.
  table: string | undefined;
  column: string;
  row: HeldRow;
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

// What a judged cell holds after the request: a value, nothing, or no cell at all.
export type CellAfter =
  | { kind: "value"; value: ExtendedValue }
  | { kind: "cleared" }
  | { kind: "deleted" }
  | { kind: "unknown"; label: string };
