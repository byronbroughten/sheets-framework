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
  "rows" | "table" | "tables"
> {
  rows: FakeCell[][];
  tables: FakeTableState[];
  rowCount: number;
  columnCount: number;
  hiddenRowIndexes: number[];
  hiddenColumnIndexes: number[];
}

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
  table(
    spreadsheet: FakeSpreadsheet,
    tableId: string | undefined,
  ): { sheet: FakeSheetState; table: FakeTableState } {
    for (const sheet of spreadsheet.sheets) {
      const table = sheet.tables.find(
        (candidate) => candidate.tableId === tableId,
      );
      if (table !== undefined) return { sheet, table };
    }
    throw new Error(`The fake spreadsheet has no Table with id ${tableId}.`);
  },
};

function sheetState(fixture: FakeSheetProperties): FakeSheetState {
  const {
    table,
    tables: fixtureTables,
    ...copy
  } = JSON.parse(JSON.stringify(fixture)) as FakeSheetProperties;
  if (table !== undefined && fixtureTables !== undefined) {
    throw new Error(
      `Fixture sheet ${copy.sheetId} gives both table and tables; give one.`,
    );
  }
  const placed = (fixtureTables ?? (table === undefined ? [] : [table])).map(
    (fixtureTable, tableIndex) =>
      placedTable(fixtureTable, defaultTableId(copy.sheetId, tableIndex)),
  );
  const rows = placed.reduce(
    withHeadRows,
    (copy.rows ?? []).map((row) => [...row]),
  );
  const widestRow = Math.max(0, ...rows.map((row) => row.length));
  // Resolved once, so a later write right of the Table never widens it.
  const tables = placed.map(({ headRows: _headRows, ...tableState }) => ({
    ...tableState,
    endColumnIndex: tableState.endColumnIndex ?? widestRow,
  }));
  return {
    ...copy,
    rows,
    tables,
    rowCount: Math.max(rows.length, ...tables.map((t) => t.endRowIndex)),
    columnCount: Math.max(widestRow, ...tables.map((t) => t.endColumnIndex)),
    hiddenRowIndexes: [],
    hiddenColumnIndexes: [],
  };
}

function defaultTableId(sheetId: number, tableIndex: number): string {
  return tableIndex === 0
    ? `fake-table-${sheetId}`
    : `fake-table-${sheetId}-${tableIndex}`;
}

type PlacedTable = FakeTable &
  Pick<FakeTableState, "tableId" | "startRowIndex" | "startColumnIndex">;

function placedTable(table: FakeTable, defaultId: string): PlacedTable {
  return {
    ...table,
    tableId: table.tableId ?? defaultId,
    startRowIndex: table.startRowIndex ?? sheetLayout.tableHeaderRowIndex,
    startColumnIndex: table.startColumnIndex ?? sheetLayout.startTableColIndex,
  };
}

// A head-row cell `rows` already holds is a fixture typo, so it throws rather than picking one.
function withHeadRows(rows: FakeCell[][], table: PlacedTable): FakeCell[][] {
  Object.entries(table.headRows ?? {}).forEach(([offsetKey, cells]) => {
    const offset = Number(offsetKey);
    const rowIndex = table.startRowIndex - offset;
    if (rowIndex < 0) {
      throw new Error(
        `Table ${table.tableId}'s head row ${offset} above its header would sit above row 0.`,
      );
    }
    while (rows.length <= rowIndex) rows.push([]);
    const row = rows[rowIndex] ?? [];
    (cells ?? []).forEach((cell, colOffset) => {
      const colIndex = table.startColumnIndex + colOffset;
      if ((row[colIndex] ?? null) !== null) {
        throw new Error(
          `Table ${table.tableId}'s head row ${offset} and rows both give a cell at row ${rowIndex}, column ${colIndex}.`,
        );
      }
      while (row.length < colIndex) row.push(null);
      row[colIndex] = cell;
    });
    rows[rowIndex] = row;
  });
  return rows;
}
