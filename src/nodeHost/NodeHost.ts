import {
  type GoogleRequest,
  GoogleSheetsAPI,
  type SheetsHttpTransport,
} from "../00_Source/GoogleSheets/GoogleSheetsAPI";
import { installRawSource } from "../00_Source/RawSource/RawSource";
import {
  type Configs,
  installConfigs,
} from "../01_SpreadsheetSchema/configRegister";
import { UpdateRequestSummary } from "./UpdateRequestSummary";

export interface NodeHostProps {
  configs: Configs;
  spreadsheetId: string;
  transport: SheetsHttpTransport;
  isDryRun: boolean;
  log: (message: string) => void;
}

// The framework's second host — Sheets only, in Node. See docs/how-it-runs.md.
export class NodeHost {
  readonly configs: Configs;
  readonly spreadsheetId: string;
  readonly isDryRun: boolean;
  private transport: SheetsHttpTransport;
  private log: (message: string) => void;
  private sentRequests: GoogleRequest[];
  constructor(props: NodeHostProps) {
    this.configs = props.configs;
    this.spreadsheetId = props.spreadsheetId;
    this.isDryRun = props.isDryRun;
    this.transport = props.transport;
    this.log = props.log;
    this.sentRequests = [];
  }
  static init(props: NodeHostProps): NodeHost {
    return new NodeHost(props);
  }
  get googleSheetsAPI(): GoogleSheetsAPI {
    return GoogleSheetsAPI.initHttp({
      spreadsheetId: this.spreadsheetId,
      transport: this.transport,
      isDryRun: this.isDryRun,
      reportRequests: (requests) => this.sentRequests.push(...requests),
    });
  }
  // Every request the run produced, sent or withheld by the dry run.
  get summary(): UpdateRequestSummary {
    return new UpdateRequestSummary({ requests: this.sentRequests });
  }
  ensureGlobals(): this {
    const globals = globalThis as Record<string, unknown>;
    globals.Logger = { log: this.log };
    installConfigs(this.configs);
    installRawSource(this.googleSheetsAPI);
    return this;
  }
}
