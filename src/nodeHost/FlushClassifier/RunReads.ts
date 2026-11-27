import { Val } from "@byronbroughten/utils/val";

import { tableLayout } from "../../01_SpreadsheetSchema/tableLayout.ts";
import { colLetters } from "../colLetters.ts";
import { ExtVal } from "../ExtVal.ts";
import type {
  ExtendedValue,
  FlushRead,
  GoogleCellData,
  GoogleGridRange,
  GridCell,
  GridSize,
  HeldCell,
  HeldColumn,
  HeldRow,
} from "./flushClassifierTypes.ts";
import { Grid } from "./Grid.ts";

type GoogleSheet = GoogleAppsScript.Sheets.Schema.Sheet;
type GoogleTable = GoogleAppsScript.Sheets.Schema.Table;
type GoogleGridProperties = GoogleAppsScript.Sheets.Schema.GridProperties;

interface SheetTables {
  tables: GoogleTable[];
  gridSize: GridSize | undefined;
  title: string | undefined;
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
  // A read of the whole tab, as fetchSheetUsedGrid asks for, leaves out only blank cells.
  hasUsedGrid(sheetGid: number): boolean {
    return (
      this.sheetTables.has(sheetGid) &&
      this.latestFirst.some(({ gridRanges }) =>
        gridRanges.some((range) => isWholeTab(range, sheetGid)),
      )
    );
  }
  isInHeadRows(cell: GridCell): boolean {
    const tables = this.sheetTables.get(cell.sheetGid)?.tables ?? [];
    return tables.some((table) => Grid.isInRange(headRowsRange(table), cell));
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
      table: tableLabel(table),
      column: this._columnName(table, cell.colIndex),
      row: this._heldRow(table, cell),
    };
  }
  // Undefined for a cell no read of its Table covered, whose column is only a letter, or whose column-ID cell reads blank; an unread column-ID cell keeps the column, for the send's batchGet to check.
  heldColumn({
    sheetGid,
    table: label,
    column,
  }: HeldCell): HeldColumn | undefined {
    const sheet = this.sheetTables.get(sheetGid);
    if (label === undefined || sheet === undefined) return undefined;
    const { title, gridSize } = sheet;
    const table = sheet.tables.find((each) => tableLabel(each) === label);
    if (title === undefined || gridSize === undefined || table === undefined) {
      return undefined;
    }
    const colIndex = this._colIndexNamed(table, column);
    if (colIndex === undefined) return undefined;
    const headerRowIndex = table.range?.startRowIndex ?? 0;
    const columnIdRowIndex =
      headerRowIndex - tableLayout.headRowOffsets.columnId;
    if (this._isReadBlank({ sheetGid, rowIndex: columnIdRowIndex, colIndex })) {
      return undefined;
    }
    return {
      sheetGid,
      tabTitle: title,
      table: label,
      tableId: Val.assert(table.tableId, "held column's Table ID"),
      header: column,
      colIndex,
      columnIdRowIndex,
      idColIndex: this._colIndexNamed(table, tableLayout.idHeader),
      tabColumnCount: gridSize.columnCount,
    };
  }
  private _isReadBlank(cell: GridCell): boolean {
    const read = this.valueAt(cell);
    return read.kind === "read" && !ExtVal.text(read.value);
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
    const colIndex = this._colIndexNamed(table, header);
    if (colIndex === undefined) return undefined;
    const data = this._cellData({ ...cell, colIndex });
    return ExtVal.text(data && readValue(data));
  }
  private _colIndexNamed(
    table: GoogleTable,
    header: string,
  ): number | undefined {
    return tableColIndexes(table).find(
      (index) => this._columnName(table, index) === header,
    );
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
        title: sheet.properties?.title,
      });
      return bySheet;
    }, new Map<number, SheetTables>());
}

function tableLabel(table: GoogleTable): string {
  return table.name ?? table.tableId ?? "(no name)";
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

function isWholeTab(range: GoogleGridRange, sheetGid: number): boolean {
  return (
    (range.sheetId ?? 0) === sheetGid &&
    range.startRowIndex === undefined &&
    range.endRowIndex === undefined &&
    range.startColumnIndex === undefined &&
    range.endColumnIndex === undefined
  );
}

function headRowsRange(table: GoogleTable): GoogleGridRange {
  const range = table.range ?? {};
  const headerRowIndex = range.startRowIndex ?? 0;
  return {
    ...range,
    startRowIndex: headerRowIndex - (tableLayout.headerZoneDepth - 1),
    endRowIndex: headerRowIndex + 1,
  };
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
