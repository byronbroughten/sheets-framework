import type {
  GoogleBatchUpdateResponse,
  SheetsHttpTransport,
} from "../GoogleSheetsAPI.ts";
import {
  postBatchUpdate,
  send,
  spreadsheetUrl,
} from "./httpSheetsTransport.ts";

// A cell as values:batchGet gives it under FORMULA: a formula's text, else its unformatted value.
export type SheetValue = string | number | boolean;

// A values:batchGet reply under majorDimension=COLUMNS: each range's columns.
interface ColumnValueRanges {
  valueRanges?: { values?: SheetValue[][] }[];
}

export interface HttpSpreadsheetProps {
  spreadsheetId: string;
  transport: SheetsHttpTransport;
}

export interface TableRange {
  startRowIndex: number;
  endRowIndex: number;
  startColumnIndex: number;
  endColumnIndex: number;
}

// A Table's ID and its range, header row first.
export interface TabTable {
  tableId: string;
  range: TableRange;
}

// A tab's ID, title, grid size and Tables.
export interface TabGrid {
  sheetGid: number;
  title: string;
  rowCount: number;
  columnCount: number;
  tables: TabTable[];
}

interface TabGridsReply {
  sheets?: GoogleAppsScript.Sheets.Schema.Sheet[];
}

export class HttpSpreadsheet {
  private spreadsheetId: string;
  private transport: SheetsHttpTransport;
  constructor({ spreadsheetId, transport }: HttpSpreadsheetProps) {
    this.spreadsheetId = spreadsheetId;
    this.transport = transport;
  }
  static init(props: HttpSpreadsheetProps): HttpSpreadsheet {
    return new HttpSpreadsheet(props);
  }
  // One values:batchGet, each range's column from its first row; Google trims a column's trailing blanks.
  columnValues(a1Ranges: string[]): SheetValue[][] {
    return this.rangeColumns(a1Ranges).map((columns) => columns[0] ?? []);
  }
  // One values:batchGet, each range as its columns; a range's trailing blank columns and each column's trailing blanks are trimmed.
  rangeColumns(a1Ranges: string[]): SheetValue[][][] {
    const query = [
      ...a1Ranges.map((range) => `ranges=${encodeURIComponent(range)}`),
      "majorDimension=COLUMNS",
      "valueRenderOption=FORMULA",
      "dateTimeRenderOption=SERIAL_NUMBER",
    ].join("&");
    const { valueRanges = [] } = send<ColumnValueRanges>(this.transport, {
      method: "GET",
      url: `${spreadsheetUrl(this.spreadsheetId, "/values:batchGet")}?${query}`,
      body: null,
    });
    return a1Ranges.map((_, index) => valueRanges[index]?.values ?? []);
  }
  // One GET of every tab's ID, title, grid size and Tables.
  tabGrids(): TabGrid[] {
    const fields =
      "sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)),tables(tableId,range))";
    const { sheets = [] } = send<TabGridsReply>(this.transport, {
      method: "GET",
      url: spreadsheetUrl(this.spreadsheetId, "", fields),
      body: null,
    });
    return sheets.map(({ properties, tables = [] }) => ({
      sheetGid: properties?.sheetId ?? 0,
      title: properties?.title ?? "",
      rowCount: properties?.gridProperties?.rowCount ?? 0,
      columnCount: properties?.gridProperties?.columnCount ?? 0,
      tables: tables.map(({ tableId, range }) => ({
        tableId: tableId ?? "",
        range: {
          startRowIndex: range?.startRowIndex ?? 0,
          endRowIndex: range?.endRowIndex ?? 0,
          startColumnIndex: range?.startColumnIndex ?? 0,
          endColumnIndex: range?.endColumnIndex ?? 0,
        },
      })),
    }));
  }
  // One batchUpdate, past the host's flush gate: for a caller that is itself the undo of a gated send.
  batchUpdate(
    requests: GoogleAppsScript.Sheets.Schema.Request[],
  ): GoogleBatchUpdateResponse {
    return postBatchUpdate(this.transport, this.spreadsheetId, { requests });
  }
}
