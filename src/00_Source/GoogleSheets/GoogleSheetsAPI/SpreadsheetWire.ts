import type {
  GoogleBatchUpdateResponse,
  SheetsHttpRequest,
  SheetsHttpTransport,
} from "../GoogleSheetsAPI.ts";

const sheetsApiBase = "https://sheets.googleapis.com/v4/spreadsheets";

export interface SpreadsheetWireProps {
  spreadsheetId: string;
  transport: SheetsHttpTransport;
}

export interface SpreadsheetUrlProps {
  suffix?: string;
  fields?: string;
}

// One spreadsheet's wire: its URLs, and the requests sent to the transport.
export class SpreadsheetWire {
  private spreadsheetId: string;
  private transport: SheetsHttpTransport;
  constructor({ spreadsheetId, transport }: SpreadsheetWireProps) {
    this.spreadsheetId = spreadsheetId;
    this.transport = transport;
  }
  static init(props: SpreadsheetWireProps): SpreadsheetWire {
    return new SpreadsheetWire(props);
  }
  url({ suffix = "", fields }: SpreadsheetUrlProps = {}): string {
    const query = fields ? `?fields=${encodeURIComponent(fields)}` : "";
    return `${sheetsApiBase}/${this.spreadsheetId}${suffix}${query}`;
  }
  postBatchUpdate(
    resource: GoogleAppsScript.Sheets.Schema.BatchUpdateSpreadsheetRequest,
  ): GoogleBatchUpdateResponse {
    return this.send<GoogleBatchUpdateResponse>({
      method: "POST",
      url: this.url({ suffix: ":batchUpdate" }),
      body: JSON.stringify(resource),
    });
  }
  // The one place the wire is trusted, as Apps Script's own declaration trusts it.
  send<RS>(request: SheetsHttpRequest): RS {
    return this.transport(request) as RS;
  }
}
