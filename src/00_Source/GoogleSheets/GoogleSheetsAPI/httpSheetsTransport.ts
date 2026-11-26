import type {
  GoogleBatchUpdateResponse,
  GoogleSheetsAPIHttpProps,
  SheetsAdvancedTransport,
  SheetsHttpRequest,
  SheetsHttpTransport,
} from "../GoogleSheetsAPI.ts";

type GoogleSpreadsheet = GoogleAppsScript.Sheets.Schema.Spreadsheet;

const sheetsApiBase = "https://sheets.googleapis.com/v4/spreadsheets";

// A cell as values:batchGet gives it under FORMULA: a formula's text, else its unformatted value.
export type SheetValue = string | number | boolean;

// A values:batchGet reply under majorDimension=COLUMNS: each range's columns.
interface ColumnValueRanges {
  valueRanges?: { values?: SheetValue[][] }[];
}

export interface HttpSpreadsheetCall {
  spreadsheetId: string;
  transport: SheetsHttpTransport;
}

export interface ColumnValuesRead extends HttpSpreadsheetCall {
  // Each an A1 range, such as 'Payments'!C:C.
  ranges: string[];
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

export interface BatchUpdateSend extends HttpSpreadsheetCall {
  requests: GoogleAppsScript.Sheets.Schema.Request[];
}

// One values:batchGet, each range's column from its first row; Google trims a column's trailing blanks.
export function httpColumnValues(read: ColumnValuesRead): SheetValue[][] {
  return httpRangeColumns(read).map((columns) => columns[0] ?? []);
}

// One values:batchGet, each range as its columns; a range's trailing blank columns and each column's trailing blanks are trimmed.
export function httpRangeColumns({
  spreadsheetId,
  transport,
  ranges,
}: ColumnValuesRead): SheetValue[][][] {
  const query = [
    ...ranges.map((range) => `ranges=${encodeURIComponent(range)}`),
    "majorDimension=COLUMNS",
    "valueRenderOption=FORMULA",
    "dateTimeRenderOption=SERIAL_NUMBER",
  ].join("&");
  const { valueRanges = [] } = send<ColumnValueRanges>(transport, {
    method: "GET",
    url: `${url(spreadsheetId, "/values:batchGet")}?${query}`,
    body: null,
  });
  return ranges.map((_, index) => valueRanges[index]?.values ?? []);
}

// One GET of every tab's ID, title, grid size and Tables.
export function httpTabGrids({
  spreadsheetId,
  transport,
}: HttpSpreadsheetCall): TabGrid[] {
  const fields =
    "sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)),tables(tableId,range))";
  const { sheets = [] } = send<TabGridsReply>(transport, {
    method: "GET",
    url: url(spreadsheetId, "", fields),
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
export function httpBatchUpdate({
  spreadsheetId,
  transport,
  requests,
}: BatchUpdateSend): GoogleBatchUpdateResponse {
  return postBatchUpdate(transport, spreadsheetId, { requests });
}

export function httpSheetsTransport(
  props: GoogleSheetsAPIHttpProps,
): SheetsAdvancedTransport {
  return {
    Spreadsheets: {
      get: (spreadsheetId, optionalArgs) => {
        const spreadsheet = send<GoogleSpreadsheet>(props.transport, {
          method: "GET",
          url: url(spreadsheetId, "", optionalArgs?.fields),
          body: null,
        });
        props.reportRead({ gridRanges: [], spreadsheet });
        return spreadsheet;
      },
      getByDataFilter: (resource, spreadsheetId, optionalArgs) => {
        const spreadsheet = send<GoogleSpreadsheet>(props.transport, {
          method: "POST",
          url: url(spreadsheetId, ":getByDataFilter", optionalArgs?.fields),
          body: JSON.stringify(resource),
        });
        props.reportRead({
          gridRanges: (resource.dataFilters ?? []).flatMap(({ gridRange }) =>
            gridRange === undefined ? [] : [gridRange],
          ),
          spreadsheet,
        });
        return spreadsheet;
      },
      batchUpdate: (resource, spreadsheetId) =>
        props.gateFlush(resource.requests ?? [], () =>
          postBatchUpdate(props.transport, spreadsheetId, resource),
        ),
    },
  };
}

function url(spreadsheetId: string, suffix: string, fields?: string): string {
  const query = fields ? `?fields=${encodeURIComponent(fields)}` : "";
  return `${sheetsApiBase}/${spreadsheetId}${suffix}${query}`;
}

function postBatchUpdate(
  transport: SheetsHttpTransport,
  spreadsheetId: string,
  resource: GoogleAppsScript.Sheets.Schema.BatchUpdateSpreadsheetRequest,
): GoogleBatchUpdateResponse {
  return send<GoogleBatchUpdateResponse>(transport, {
    method: "POST",
    url: url(spreadsheetId, ":batchUpdate"),
    body: JSON.stringify(resource),
  });
}

// The one place the wire is trusted, as Apps Script's own declaration trusts it.
function send<RS>(
  transport: SheetsHttpTransport,
  request: SheetsHttpRequest,
): RS {
  return transport(request) as RS;
}
