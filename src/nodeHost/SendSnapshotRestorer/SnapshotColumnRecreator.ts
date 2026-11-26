import type {
  SheetValue,
  TabTable,
} from "../../00_Source/GoogleSheets/GoogleSheetsAPI/httpSheetsTransport.ts";
import { tableLayout } from "../../01_SpreadsheetSchema/tableLayout.ts";
import type { SnapshotColumn } from "../SendSnapshotTaker.ts";
import { type CurrentTab, enteredCell } from "./SnapshotColumnRestorer.ts";

type GoogleRequest = GoogleAppsScript.Sheets.Schema.Request;

export interface SnapshotColumnRecreatorProps {
  tab: CurrentTab;
  // The tab's saved columns whose column ID is gone from it.
  columns: SnapshotColumn[];
}

// The tab as the recreations leave it, so a later column-ID lookup finds each recreated column.
export interface ColumnRecreation {
  tab: CurrentTab;
  requests: GoogleRequest[];
}

interface ColumnBand {
  startRowIndex: number;
  endRowIndex: number;
  colIndex: number;
}

const headerRowOffset =
  tableLayout.headRowOffsets.columnId - tableLayout.headRowOffsets.header;

// Each deleted column back at its saved position, head rows and column ID included, as the framework inserts one: from the column ID row to the Table's last row.
export class SnapshotColumnRecreator {
  private tab: CurrentTab;
  private columns: SnapshotColumn[];
  constructor(props: SnapshotColumnRecreatorProps) {
    this.tab = props.tab;
    this.columns = props.columns;
  }
  static init(props: SnapshotColumnRecreatorProps): SnapshotColumnRecreator {
    return new SnapshotColumnRecreator(props);
  }
  // Leftmost first, so each saved position is read in the pre-send layout the ones before it rebuild.
  recreate(): ColumnRecreation {
    const sorted = this.columns.toSorted((a, b) => a.colIndex - b.colIndex);
    return sorted.reduce(
      (recreation, column) => recreatedOn(recreation, column),
      { tab: this.tab, requests: [] } as ColumnRecreation,
    );
  }
}

// A column whose Table is gone is left as it is, for the restore to list as not restored.
function recreatedOn(
  { tab, requests }: ColumnRecreation,
  column: SnapshotColumn,
): ColumnRecreation {
  const table = tableOf(tab, column);
  if (table === undefined) return { tab, requests };
  const band: ColumnBand = {
    startRowIndex: column.columnIdRowIndex,
    endRowIndex: table.endRowIndex,
    colIndex: column.colIndex,
  };
  const headValues = column.values.slice(
    column.columnIdRowIndex,
    column.columnIdRowIndex + tableLayout.headerZoneDepth,
  );
  return {
    tab: {
      grid: { ...tab.grid, tables: tablesAfter(tab.grid.tables, table, band) },
      columns: columnsAfter(tab.columns, band, headValues),
    },
    requests: [
      ...requests,
      ...insertRequests(column.sheetGid, table, band),
      headRowWrite(column.sheetGid, band, headValues),
    ],
  };
}

// The Table the column stood in: its header row under the column's column ID row, and the column inside it or at an edge.
function tableOf(
  tab: CurrentTab,
  { columnIdRowIndex, colIndex }: SnapshotColumn,
): TabTable | undefined {
  return tab.grid.tables.find(
    (table) =>
      table.startRowIndex === columnIdRowIndex + headerRowOffset &&
      table.startColumnIndex <= colIndex &&
      colIndex <= table.endColumnIndex,
  );
}

// An insert at a Table's edge isn't counted on to widen it, so the widen is sent too, as the framework's end-column insert sends one.
function insertRequests(
  sheetId: number,
  table: TabTable,
  band: ColumnBand,
): GoogleRequest[] {
  const insert: GoogleRequest = {
    insertRange: {
      range: {
        sheetId,
        startRowIndex: band.startRowIndex,
        endRowIndex: band.endRowIndex,
        startColumnIndex: band.colIndex,
        endColumnIndex: band.colIndex + 1,
      },
      shiftDimension: "COLUMNS",
    },
  };
  const isInside =
    table.startColumnIndex < band.colIndex &&
    band.colIndex < table.endColumnIndex;
  if (isInside) return [insert];
  const widened = widenedTable(table, band);
  return [
    insert,
    {
      updateTable: {
        table: {
          tableId: table.tableId,
          range: {
            sheetId,
            startRowIndex: widened.startRowIndex,
            endRowIndex: widened.endRowIndex,
            startColumnIndex: widened.startColumnIndex,
            endColumnIndex: widened.endColumnIndex,
          },
        },
        fields: "range",
      },
    },
  ];
}

function widenedTable(table: TabTable, { colIndex }: ColumnBand): TabTable {
  return {
    ...table,
    startColumnIndex: Math.min(table.startColumnIndex, colIndex),
    endColumnIndex: table.endColumnIndex + 1,
  };
}

function headRowWrite(
  sheetId: number,
  band: ColumnBand,
  headValues: SheetValue[],
): GoogleRequest {
  return {
    updateCells: {
      range: {
        sheetId,
        startRowIndex: band.startRowIndex,
        endRowIndex: band.startRowIndex + tableLayout.headerZoneDepth,
        startColumnIndex: band.colIndex,
        endColumnIndex: band.colIndex + 1,
      },
      rows: Array.from({ length: tableLayout.headerZoneDepth }, (_, index) => ({
        values: [enteredCell(headValues[index] ?? "")],
      })),
      fields: "userEnteredValue",
    },
  };
}

// The Table widened, and a Table wholly in the band's rows and right of it shifted right, as the insert leaves them.
function tablesAfter(
  tables: TabTable[],
  recreatedIn: TabTable,
  band: ColumnBand,
): TabTable[] {
  return tables.map((table) => {
    if (table === recreatedIn) return widenedTable(table, band);
    const isShifted =
      table.startColumnIndex >= band.colIndex &&
      band.startRowIndex <= table.startRowIndex &&
      table.endRowIndex <= band.endRowIndex;
    if (!isShifted) return table;
    return {
      ...table,
      startColumnIndex: table.startColumnIndex + 1,
      endColumnIndex: table.endColumnIndex + 1,
    };
  });
}

// The band's rows shifted one column right from the column, which then holds its saved head rows.
function columnsAfter(
  columns: SheetValue[][],
  band: ColumnBand,
  headValues: SheetValue[],
): SheetValue[][] {
  const width = Math.max(columns.length, band.colIndex) + 1;
  const after = Array.from({ length: width }, (_, colIndex) => [
    ...(columns[colIndex] ?? []),
  ]);
  for (
    let rowIndex = band.startRowIndex;
    rowIndex < band.endRowIndex;
    rowIndex++
  ) {
    for (let colIndex = width - 1; colIndex > band.colIndex; colIndex--) {
      const shifted = after[colIndex];
      if (shifted !== undefined) {
        shifted[rowIndex] = after[colIndex - 1]?.[rowIndex] ?? "";
      }
    }
    const recreated = after[band.colIndex];
    if (recreated !== undefined) {
      recreated[rowIndex] = headValues[rowIndex - band.startRowIndex] ?? "";
    }
  }
  return after;
}
