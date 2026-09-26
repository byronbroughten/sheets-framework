import type { FakeCell, FakeCellValue } from "../fakeSheetsService";
import { fakeCells } from "./fakeCells";
import { fakeGrid } from "./fakeGrid";
import {
  type FakeSheetState,
  type FakeSpreadsheet,
  fakeSpreadsheet,
} from "./fakeSpreadsheet";
import { fakeTables } from "./fakeTables";

export interface FakeGridRange {
  startRowIndex?: number;
  endRowIndex?: number;
  startColumnIndex?: number;
  endColumnIndex?: number;
}

// One sheet as the fake spreadsheet holds it when `sheet()` is called; a copy, so a test can't edit it.
export interface FakeSheetView {
  title: string;
  rowCount: number;
  columnCount: number;
  hiddenRowIndexes: number[];
  hiddenColumnIndexes: number[];
  tables: GoogleAppsScript.Sheets.Schema.Table[];
  protectedRanges: GoogleAppsScript.Sheets.Schema.ProtectedRange[];
  conditionalFormats: GoogleAppsScript.Sheets.Schema.ConditionalFormatRule[];
  cell(rowIndex: number, colIndex: number): FakeCell;
  // A rectangle of cells, `null` where empty; with no range, row 0 and column 0 to the last cell holding anything.
  rows(range?: FakeGridRange): FakeCell[][];
  values(range?: FakeGridRange): FakeCellValue[][];
}

export interface FakeGridView {
  sheetTitles(): string[];
  sheet(sheetGid: number): FakeSheetView;
}

export function buildGridView(spreadsheet: FakeSpreadsheet): FakeGridView {
  return {
    sheetTitles() {
      return spreadsheet.sheets.map((sheet) => sheet.title);
    },
    sheet(sheetGid) {
      return sheetView(fakeSpreadsheet.sheet(spreadsheet, sheetGid));
    },
  };
}

function sheetView(state: FakeSheetState): FakeSheetView {
  const sheet = copied(state);
  function rows(range: FakeGridRange = {}): FakeCell[][] {
    const bounded = {
      startRowIndex: range.startRowIndex ?? 0,
      endRowIndex: range.endRowIndex ?? filledEndRowIndex(sheet),
      startColumnIndex: range.startColumnIndex ?? 0,
      endColumnIndex: range.endColumnIndex ?? filledEndColumnIndex(sheet),
    };
    return Array.from(
      { length: bounded.endRowIndex - bounded.startRowIndex },
      (_, rowOffset) =>
        Array.from(
          { length: bounded.endColumnIndex - bounded.startColumnIndex },
          (_unused, colOffset) =>
            fakeGrid.cell(
              sheet,
              bounded.startRowIndex + rowOffset,
              bounded.startColumnIndex + colOffset,
            ),
        ),
    );
  }
  return {
    title: sheet.title,
    rowCount: sheet.rowCount,
    columnCount: sheet.columnCount,
    hiddenRowIndexes: sheet.hiddenRowIndexes,
    hiddenColumnIndexes: sheet.hiddenColumnIndexes,
    tables: fakeTables.googleTables(sheet, false) ?? [],
    protectedRanges: sheet.protectedRanges ?? [],
    conditionalFormats: sheet.conditionalFormats ?? [],
    cell(rowIndex, colIndex) {
      return fakeGrid.cell(sheet, rowIndex, colIndex);
    },
    rows,
    values(range) {
      return rows(range).map((row) => row.map(fakeCells.value));
    },
  };
}

function copied(sheet: FakeSheetState): FakeSheetState {
  return JSON.parse(JSON.stringify(sheet)) as FakeSheetState;
}

function filledEndRowIndex(sheet: FakeSheetState): number {
  return (
    sheet.rows.findLastIndex((row) => row.some((cell) => cell !== null)) + 1
  );
}

function filledEndColumnIndex(sheet: FakeSheetState): number {
  return Math.max(
    0,
    ...sheet.rows.map((row) => row.findLastIndex((cell) => cell !== null) + 1),
  );
}
