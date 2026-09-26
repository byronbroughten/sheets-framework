import { sheetLayout } from "../../01_SpreadsheetSchema/sheetLayout";
import type {
  FakeCell,
  FakeSheetProperties,
  FakeTable,
} from "../fakeSheetsService";

export interface FakeTableState extends FakeTable {
  tableId: string;
  startRowIndex: number;
  startColumnIndex: number;
  endColumnIndex: number;
}

export interface FakeSheetState extends Omit<
  FakeSheetProperties,
  "rows" | "table" | "extraTables"
> {
  rows: FakeCell[][];
  table?: FakeTableState;
  extraTables: FakeTableState[];
  rowCount: number;
  columnCount: number;
  hiddenRowIndexes: number[];
  hiddenColumnIndexes: number[];
}

// The in-memory spreadsheet every read serves and every batch update replays onto.
export interface FakeSpreadsheet {
  sheets: FakeSheetState[];
  lastProtectedRangeId: number;
}

export const fakeSpreadsheet = {
  // A copy, so a fixture shared across tests never carries one test's writes into the next.
  init(fixtures: readonly FakeSheetProperties[]): FakeSpreadsheet {
    const sheets = fixtures.map(sheetState);
    const lastProtectedRangeId = Math.max(
      0,
      ...sheets.flatMap((sheet) =>
        (sheet.protectedRanges ?? []).map(
          (protection) => protection.protectedRangeId ?? 0,
        ),
      ),
    );
    return { sheets, lastProtectedRangeId };
  },
  sheet(
    spreadsheet: FakeSpreadsheet,
    sheetId: number | undefined,
  ): FakeSheetState {
    const sheet = spreadsheet.sheets.find(
      (candidate) => candidate.sheetId === sheetId,
    );
    if (sheet === undefined) {
      throw new Error(`The fake spreadsheet has no sheet with gid ${sheetId}.`);
    }
    return sheet;
  },
  tables(sheet: FakeSheetState): FakeTableState[] {
    return sheet.table === undefined
      ? [...sheet.extraTables]
      : [sheet.table, ...sheet.extraTables];
  },
  table(
    spreadsheet: FakeSpreadsheet,
    tableId: string | undefined,
  ): { sheet: FakeSheetState; table: FakeTableState } {
    for (const sheet of spreadsheet.sheets) {
      const table = fakeSpreadsheet
        .tables(sheet)
        .find((candidate) => candidate.tableId === tableId);
      if (table !== undefined) return { sheet, table };
    }
    throw new Error(`The fake spreadsheet has no Table with id ${tableId}.`);
  },
};

function sheetState(fixture: FakeSheetProperties): FakeSheetState {
  const copy = JSON.parse(JSON.stringify(fixture)) as FakeSheetProperties;
  const rows = (copy.rows ?? []).map((row) => [...row]);
  const widestRow = Math.max(0, ...rows.map((row) => row.length));
  const table =
    copy.table === undefined
      ? undefined
      : tableState(copy.table, `fake-table-${copy.sheetId}`, widestRow);
  const extraTables = (copy.extraTables ?? []).map((extraTable, extraIndex) =>
    tableState(
      extraTable,
      `fake-table-${copy.sheetId}-extra-${extraIndex}`,
      widestRow,
    ),
  );
  const tables = table === undefined ? extraTables : [table, ...extraTables];
  return {
    ...copy,
    rows,
    table,
    extraTables,
    rowCount: Math.max(rows.length, ...tables.map((t) => t.endRowIndex)),
    columnCount: Math.max(widestRow, ...tables.map((t) => t.endColumnIndex)),
    hiddenRowIndexes: [],
    hiddenColumnIndexes: [],
  };
}

// Resolved once, so a later write right of the Table never widens it.
function tableState(
  table: FakeTable,
  defaultTableId: string,
  widestRow: number,
): FakeTableState {
  return {
    ...table,
    tableId: table.tableId ?? defaultTableId,
    startRowIndex: table.startRowIndex ?? sheetLayout.tableHeaderRowIndex,
    startColumnIndex: table.startColumnIndex ?? sheetLayout.startTableColIndex,
    endColumnIndex: table.endColumnIndex ?? widestRow,
  };
}
