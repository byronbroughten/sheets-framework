import type {
  GoogleBatchUpdateResponse,
  GoogleSheetsAPIHttpProps,
  SheetsAdvancedTransport,
  SheetsHttpRequest,
  SheetsHttpTransport,
} from "../GoogleSheetsAPI.ts";

type GoogleSpreadsheet = GoogleAppsScript.Sheets.Schema.Spreadsheet;

const sheetsApiBase = "https://sheets.googleapis.com/v4/spreadsheets";

export function httpSheetsTransport(
  props: GoogleSheetsAPIHttpProps,
): SheetsAdvancedTransport {
  return {
    Spreadsheets: {
      get: (spreadsheetId, optionalArgs) => {
        const spreadsheet = send<GoogleSpreadsheet>(props.transport, {
          method: "GET",
          url: spreadsheetUrl(spreadsheetId, "", optionalArgs?.fields),
          body: null,
        });
        props.reportRead({ gridRanges: [], spreadsheet });
        return spreadsheet;
      },
      getByDataFilter: (resource, spreadsheetId, optionalArgs) => {
        const spreadsheet = send<GoogleSpreadsheet>(props.transport, {
          method: "POST",
          url: spreadsheetUrl(
            spreadsheetId,
            ":getByDataFilter",
            optionalArgs?.fields,
          ),
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

export function spreadsheetUrl(
  spreadsheetId: string,
  suffix: string,
  fields?: string,
): string {
  const query = fields ? `?fields=${encodeURIComponent(fields)}` : "";
  return `${sheetsApiBase}/${spreadsheetId}${suffix}${query}`;
}

export function postBatchUpdate(
  transport: SheetsHttpTransport,
  spreadsheetId: string,
  resource: GoogleAppsScript.Sheets.Schema.BatchUpdateSpreadsheetRequest,
): GoogleBatchUpdateResponse {
  return send<GoogleBatchUpdateResponse>(transport, {
    method: "POST",
    url: spreadsheetUrl(spreadsheetId, ":batchUpdate"),
    body: JSON.stringify(resource),
  });
}

// The one place the wire is trusted, as Apps Script's own declaration trusts it.
export function send<RS>(
  transport: SheetsHttpTransport,
  request: SheetsHttpRequest,
): RS {
  return transport(request) as RS;
}
