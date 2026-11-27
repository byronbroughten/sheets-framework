import type {
  SheetValue,
  TableRange,
  TabTable,
} from "../../00_Source/GoogleSheets/GoogleSheetsAPI/HttpSpreadsheet.ts";
import { tableLayout } from "../../01_SpreadsheetSchema/tableLayout.ts";
import type { SnapshotColumn } from "../SendSnapshotTaker.ts";
import { type CurrentTab, enteredCell } from "./SnapshotColumnRestorer.ts";

type GoogleRequest = GoogleAppsScript.Sheets.Schema.Request;

export interface RecreatedColumnProps {
  column: SnapshotColumn;
  // The Table the column stood in.
  table: TabTable;
}

interface ColumnBand {
  startRowIndex: number;
  endRowIndex: number;
  colIndex: number;
}

// One deleted column back at its saved position, as the framework inserts one: from the column ID row to the Table's last row.
export class RecreatedColumn {
  private sheetGid: number;
  private table: TabTable;
  private band: ColumnBand;
  private headValues: SheetValue[];
  constructor({ column, table }: RecreatedColumnProps) {
    this.sheetGid = column.sheetGid;
    this.table = table;
    this.band = {
      startRowIndex: column.columnIdRowIndex,
      endRowIndex: table.range.endRowIndex,
      colIndex: column.colIndex,
    };
    this.headValues = column.values.slice(
      column.columnIdRowIndex,
      column.columnIdRowIndex + tableLayout.headerZoneDepth,
    );
  }
  static init(props: RecreatedColumnProps): RecreatedColumn {
    return new RecreatedColumn(props);
  }
  get requests(): GoogleRequest[] {
    return [...this._insertRequests, this._headRowWrite];
  }
  // The tab as the insert and head-row write leave it.
  tabAfter(tab: CurrentTab): CurrentTab {
    return {
      grid: { ...tab.grid, tables: this._tablesAfter(tab.grid.tables) },
      columns: this._columnsAfter(tab.columns),
    };
  }
  // An insert at a Table's edge isn't counted on to widen it, so the widen is sent too, as the framework's end-column insert sends one.
  private get _insertRequests(): GoogleRequest[] {
    const { startRowIndex, endRowIndex, colIndex } = this.band;
    const insert: GoogleRequest = {
      insertRange: {
        range: {
          sheetId: this.sheetGid,
          startRowIndex,
          endRowIndex,
          startColumnIndex: colIndex,
          endColumnIndex: colIndex + 1,
        },
        shiftDimension: "COLUMNS",
      },
    };
    const { range } = this.table;
    const isInside =
      range.startColumnIndex < colIndex && colIndex < range.endColumnIndex;
    if (isInside) return [insert];
    return [
      insert,
      {
        updateTable: {
          table: {
            tableId: this.table.tableId,
            range: { sheetId: this.sheetGid, ...this._widenedRange },
          },
          fields: "range",
        },
      },
    ];
  }
  private get _widenedRange(): TableRange {
    const { range } = this.table;
    return {
      ...range,
      startColumnIndex: Math.min(range.startColumnIndex, this.band.colIndex),
      endColumnIndex: range.endColumnIndex + 1,
    };
  }
  private get _headRowWrite(): GoogleRequest {
    const { startRowIndex, colIndex } = this.band;
    return {
      updateCells: {
        range: {
          sheetId: this.sheetGid,
          startRowIndex,
          endRowIndex: startRowIndex + tableLayout.headerZoneDepth,
          startColumnIndex: colIndex,
          endColumnIndex: colIndex + 1,
        },
        rows: Array.from(
          { length: tableLayout.headerZoneDepth },
          (_, index) => ({
            values: [enteredCell(this.headValues[index] ?? "")],
          }),
        ),
        fields: "userEnteredValue",
      },
    };
  }
  // The Table widened, and a Table wholly in the band's rows and right of it shifted right, as the insert leaves them.
  private _tablesAfter(tables: TabTable[]): TabTable[] {
    const { startRowIndex, endRowIndex, colIndex } = this.band;
    return tables.map((table) => {
      if (table === this.table) return { ...table, range: this._widenedRange };
      const { range } = table;
      const isShifted =
        range.startColumnIndex >= colIndex &&
        startRowIndex <= range.startRowIndex &&
        range.endRowIndex <= endRowIndex;
      if (!isShifted) return table;
      return {
        ...table,
        range: {
          ...range,
          startColumnIndex: range.startColumnIndex + 1,
          endColumnIndex: range.endColumnIndex + 1,
        },
      };
    });
  }
  // The band's rows shifted one column right from the column, which then holds its saved head rows.
  private _columnsAfter(columns: SheetValue[][]): SheetValue[][] {
    const { startRowIndex, endRowIndex, colIndex: bandCol } = this.band;
    const width = Math.max(columns.length, bandCol) + 1;
    const after = Array.from({ length: width }, (_, colIndex) => [
      ...(columns[colIndex] ?? []),
    ]);
    for (let rowIndex = startRowIndex; rowIndex < endRowIndex; rowIndex++) {
      for (let colIndex = width - 1; colIndex > bandCol; colIndex--) {
        const shifted = after[colIndex];
        if (shifted !== undefined) {
          shifted[rowIndex] = after[colIndex - 1]?.[rowIndex] ?? "";
        }
      }
      const recreated = after[bandCol];
      if (recreated !== undefined) {
        recreated[rowIndex] = this.headValues[rowIndex - startRowIndex] ?? "";
      }
    }
    return after;
  }
}
