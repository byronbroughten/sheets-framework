import type {
  GoogleBatchUpdateResponse,
  GoogleSheetsAPIHttpProps,
  SheetsAdvancedTransport,
  SheetsHttpRequest,
} from "../GoogleSheetsAPI";

type GoogleSpreadsheet = GoogleAppsScript.Sheets.Schema.Spreadsheet;

const sheetsApiBase = "https://sheets.googleapis.com/v4/spreadsheets";

export function httpSheetsTransport(
  props: GoogleSheetsAPIHttpProps,
): SheetsAdvancedTransport {
  // The one place the wire is trusted, as Apps Script's own declaration trusts it.
  function send<RS>(request: SheetsHttpRequest): RS {
    return props.transport(request) as RS;
  }
  function url(spreadsheetId: string, suffix: string, fields?: string): string {
    const query = fields ? `?fields=${encodeURIComponent(fields)}` : "";
    return `${sheetsApiBase}/${spreadsheetId}${suffix}${query}`;
  }
  return {
    Spreadsheets: {
      get: (spreadsheetId, optionalArgs) => {
        const spreadsheet = send<GoogleSpreadsheet>({
          method: "GET",
          url: url(spreadsheetId, "", optionalArgs?.fields),
          body: null,
        });
        props.reportRead({ gridRanges: [], spreadsheet });
        return spreadsheet;
      },
      getByDataFilter: (resource, spreadsheetId, optionalArgs) => {
        const spreadsheet = send<GoogleSpreadsheet>({
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
          send<GoogleBatchUpdateResponse>({
            method: "POST",
            url: url(spreadsheetId, ":batchUpdate"),
            body: JSON.stringify(resource),
          }),
        ),
    },
  };
}
