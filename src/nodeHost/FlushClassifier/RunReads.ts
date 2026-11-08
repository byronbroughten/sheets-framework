import { tableLayout } from "../../01_SpreadsheetSchema/tableLayout";
import { colLetters } from "../colLetters";
import { ExtVal } from "../ExtVal";
import type {
  ExtendedValue,
  FlushRead,
  GoogleCellData,
  GridCell,
  GridSize,
  HeldCell,
  HeldRow,
} from "./flushClassifierTypes";
import { Grid } from "./Grid";

type GoogleSheet = GoogleAppsScript.Sheets.Schema.Sheet;
type GoogleTable = GoogleAppsScript.Sheets.Schema.Table;
type GoogleGridProperties = GoogleAppsScript.Sheets.Schema.GridProperties;

interface SheetTables {
  tables: GoogleTable[];
  gridSize: GridSize | undefined;
}

export type ReadValue =
  { kind: "unread" } | { kind: "read"; value: ExtendedValue | undefined };

export type ReadCell = { kind: "outsideTables" } | ReadValue;

// The cells as the run read them, before this flush: which Table holds each and what it held.
export class RunReads {
  private latestFirst: FlushRead[];
  private sheetTables: Map<number, SheetTables>;
  constructor(reads: FlushRead[]) {
    this.latestFirst = [...reads].reverse();
    this.sheetTables = tablesBySheet(reads);
  }
  get gridSizes(): Map<number, GridSize> {
    return [...this.sheetTables].reduce((sizes, [sheetGid, { gridSize }]) => {
      if (gridSize !== undefined) sizes.set(sheetGid, gridSize);
      return sizes;
    }, new Map<number, GridSize>());
  }
  table(tableId: string): GoogleTable | undefined {
    return [...this.sheetTables.values()]
      .flatMap(({ tables }) => tables)
      .find((table) => table.tableId === tableId);
  }
  cellAt(cell: GridCell): ReadCell {
    const sheet = this.sheetTables.get(cell.sheetGid);
    if (sheet === undefined) return { kind: "unread" };
    if (bodyTableAt(sheet.tables, cell) === undefined) {
      return { kind: "outsideTables" };
    }
    return this.valueAt(cell);
  }
  valueAt(cell: GridCell): ReadValue {
    const data = this._cellData(cell);
    if (data === undefined) return { kind: "unread" };
    return { kind: "read", value: readValue(data) };
  }
  heldCell(cell: GridCell): HeldCell {
    const table = bodyTableAt(
      this.sheetTables.get(cell.sheetGid)?.tables ?? [],
      cell,
    );
    if (table === undefined) {
      return {
        sheetGid: cell.sheetGid,
        table: undefined,
        column: colLetters(cell.colIndex),
        row: { sheetRow: cell.rowIndex + 1, nameValue: undefined },
      };
    }
    return {
      sheetGid: cell.sheetGid,
      table: table.name ?? table.tableId ?? "(no name)",
      column: this._columnName(table, cell.colIndex),
      row: this._heldRow(table, cell),
    };
  }
  private _heldRow(table: GoogleTable, cell: GridCell): HeldRow {
    const rowId = this._textInColumn(table, cell, tableLayout.idHeader);
    if (rowId) return { rowId };
    return {
      sheetRow: cell.rowIndex + 1,
      nameValue: this._textInColumn(table, cell, tableLayout.nameHeader),
    };
  }
  private _textInColumn(
    table: GoogleTable,
    cell: GridCell,
    header: string,
  ): string | undefined {
    const colIndex = tableColIndexes(table).find(
      (index) => this._columnName(table, index) === header,
    );
    if (colIndex === undefined) return undefined;
    const data = this._cellData({ ...cell, colIndex });
    return ExtVal.text(data && readValue(data));
  }
  private _columnName(table: GoogleTable, colIndex: number): string {
    const range = table.range ?? {};
    const tableColIndex = colIndex - (range.startColumnIndex ?? 0);
    // Google omits a zero columnIndex; observed: docs/sheets-behaviour.md#empty-lists-and-zero-valued-fields.
    const columnName = table.columnProperties?.find(
      (column) => (column.columnIndex ?? 0) === tableColIndex,
    )?.columnName;
    if (columnName !== undefined) return columnName;
    const header = this._cellData({
      sheetGid: range.sheetId ?? 0,
      rowIndex: range.startRowIndex ?? 0,
      colIndex,
    });
    return ExtVal.text(header && readValue(header)) ?? colLetters(colIndex);
  }
  // A later read wins, since it saw the sheet after any flush between them.
  private _cellData(cell: GridCell): GoogleCellData | undefined {
    const read = this.latestFirst.find(({ gridRanges }) =>
      gridRanges.some((range) => Grid.isInRange(range, cell)),
    );
    if (read === undefined) return undefined;
    return blockCellData(read, cell) ?? {};
  }
}

function tablesBySheet(reads: FlushRead[]): Map<number, SheetTables> {
  return reads
    .flatMap((read) => read.spreadsheet.sheets ?? [])
    .filter(hasTablesRead)
    .reduce((bySheet, sheet) => {
      bySheet.set(sheet.properties?.sheetId ?? 0, {
        tables: sheet.tables ?? [],
        gridSize: gridSizeOf(sheet.properties?.gridProperties),
      });
      return bySheet;
    }, new Map<number, SheetTables>());
}

function gridSizeOf(
  gridProperties: GoogleGridProperties | undefined,
): GridSize | undefined {
  const rowCount = gridProperties?.rowCount;
  const columnCount = gridProperties?.columnCount;
  if (rowCount === undefined || columnCount === undefined) return undefined;
  return { rowCount, columnCount };
}

// Google omits an empty tables list, but every read that asks for Tables asks for grid properties too.
function hasTablesRead(sheet: GoogleSheet): boolean {
  return (
    sheet.tables !== undefined || sheet.properties?.gridProperties !== undefined
  );
}

// Body rows only: the header row and the head rows above it are never entered data.
function bodyTableAt(
  tables: GoogleTable[],
  cell: GridCell,
): GoogleTable | undefined {
  return tables.find((table) => {
    const range = table.range ?? {};
    return (
      Grid.isInRange(range, cell) && cell.rowIndex > (range.startRowIndex ?? 0)
    );
  });
}

function tableColIndexes(table: GoogleTable): number[] {
  const { startColumnIndex = 0, endColumnIndex = startColumnIndex } =
    table.range ?? {};
  return Array.from(
    { length: endColumnIndex - startColumnIndex },
    (_, offset) => startColumnIndex + offset,
  );
}

function blockCellData(
  read: FlushRead,
  cell: GridCell,
): GoogleCellData | undefined {
  const sheet = read.spreadsheet.sheets?.find(
    (candidate) => (candidate.properties?.sheetId ?? 0) === cell.sheetGid,
  );
  return (sheet?.data ?? [])
    .map((block) => {
      const rowOffset = cell.rowIndex - (block.startRow ?? 0);
      const colOffset = cell.colIndex - (block.startColumn ?? 0);
      if (rowOffset < 0 || colOffset < 0) return undefined;
      return block.rowData?.[rowOffset]?.values?.[colOffset];
    })
    .find((data) => data !== undefined);
}

// A formula's text, else what the cell shows; undefined for an empty cell.
function readValue(data: GoogleCellData): ExtendedValue | undefined {
  const formulaValue = data.userEnteredValue?.formulaValue;
  if (formulaValue !== undefined) return { formulaValue };
  const shown = data.effectiveValue ?? data.userEnteredValue;
  if (shown === undefined) return undefined;
  const { stringValue, numberValue, boolValue, errorValue } = shown;
  if (
    stringValue === undefined &&
    numberValue === undefined &&
    boolValue === undefined &&
    errorValue === undefined
  ) {
    return undefined;
  }
  return shown;
}
