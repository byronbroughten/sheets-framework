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

// A values:batchGet reply under majorDimension=COLUMNS: each range's one column.
interface ColumnValueRanges {
  valueRanges?: { values?: SheetValue[][] }[];
}

export interface ColumnValuesRead {
  spreadsheetId: string;
  transport: SheetsHttpTransport;
  // Each an A1 range of one column, such as 'Payments'!C:C.
  ranges: string[];
}

// One values:batchGet, each range's column from its first row; Google trims a column's trailing blanks.
export function httpColumnValues({
  spreadsheetId,
  transport,
  ranges,
}: ColumnValuesRead): SheetValue[][] {
  const query = [
    ...ranges.map((range) => `ranges=${encodeURIComponent(range)}`),
    "majorDimension=COLUMNS",
    "valueRenderOption=FORMULA",
    "dateTimeRenderOption=SERIAL_NUMBER",
  ].join("&");
  const { valueRanges = [] } = send<ColumnValueRanges>(transport, {
    method: "GET",
    url: `${sheetsApiBase}/${spreadsheetId}/values:batchGet?${query}`,
    body: null,
  });
  return ranges.map((_, index) => valueRanges[index]?.values?.[0] ?? []);
}

export function httpSheetsTransport(
  props: GoogleSheetsAPIHttpProps,
): SheetsAdvancedTransport {
  function url(spreadsheetId: string, suffix: string, fields?: string): string {
    const query = fields ? `?fields=${encodeURIComponent(fields)}` : "";
    return `${sheetsApiBase}/${spreadsheetId}${suffix}${query}`;
  }
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
          send<GoogleBatchUpdateResponse>(props.transport, {
            method: "POST",
            url: url(spreadsheetId, ":batchUpdate"),
            body: JSON.stringify(resource),
          }),
        ),
    },
  };
}

// The one place the wire is trusted, as Apps Script's own declaration trusts it.
function send<RS>(
  transport: SheetsHttpTransport,
  request: SheetsHttpRequest,
): RS {
  return transport(request) as RS;
}
