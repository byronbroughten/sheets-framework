import type {
  GoogleSheetsAPIHttpProps,
  SheetsAdvancedTransport,
} from "../GoogleSheetsAPI.ts";
import { SpreadsheetWire } from "./SpreadsheetWire.ts";

type GoogleSpreadsheet = GoogleAppsScript.Sheets.Schema.Spreadsheet;

export function httpSheetsTransport(
  props: GoogleSheetsAPIHttpProps,
): SheetsAdvancedTransport {
  function wireOf(spreadsheetId: string): SpreadsheetWire {
    return SpreadsheetWire.init({ spreadsheetId, transport: props.transport });
  }
  return {
    Spreadsheets: {
      get: (spreadsheetId, optionalArgs) => {
        const wire = wireOf(spreadsheetId);
        const spreadsheet = wire.send<GoogleSpreadsheet>({
          method: "GET",
          url: wire.url({ fields: optionalArgs?.fields }),
          body: null,
        });
        props.reportRead({ gridRanges: [], spreadsheet });
        return spreadsheet;
      },
      getByDataFilter: (resource, spreadsheetId, optionalArgs) => {
        const wire = wireOf(spreadsheetId);
        const spreadsheet = wire.send<GoogleSpreadsheet>({
          method: "POST",
          url: wire.url({
            suffix: ":getByDataFilter",
            fields: optionalArgs?.fields,
          }),
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
          wireOf(spreadsheetId).postBatchUpdate(resource),
        ),
    },
  };
}
