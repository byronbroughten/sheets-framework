import {
  type GoogleBatchUpdateResponse,
  type GoogleRequest,
  GoogleSheetsAPI,
  type SheetsHttpTransport,
} from "../00_Source/GoogleSheets/GoogleSheetsAPI.ts";
import { installRawSource } from "../00_Source/RawSource/installedRawSource.ts";
import {
  type Configs,
  installConfigs,
} from "../01_SpreadsheetSchema/configReaders/configRegister.ts";
import type { Endpoints } from "../06_API/Endpoints.ts";
import { installEndpoints } from "../06_API/feedbackColumnIds.ts";
import {
  EnteredDataGate,
  type EnteredDataGateProps,
  type HeldList,
} from "./EnteredDataGate.ts";
import { type SendSnapshot, SendSnapshotTaker } from "./SendSnapshotTaker.ts";
import { UpdateRequestSummary } from "./UpdateRequestSummary.ts";

export interface NodeHostProps {
  configs: Configs;
  // The app's map; absent, only the framework's own endpoints are installed.
  endpoints?: Endpoints;
  spreadsheetId: string;
  transport: SheetsHttpTransport;
  isDryRun: boolean;
  log: (message: string) => void;
  // Sync and chores pass one; without it, every flush goes out unclassified.
  enteredDataGate?: EnteredDataGateProps;
}

// The framework's second host — Sheets only, in Node. See docs/how-it-runs/node-host.md.
export class NodeHost {
  readonly configs: Configs;
  private endpoints: Endpoints;
  readonly spreadsheetId: string;
  readonly isDryRun: boolean;
  private enteredDataGate: EnteredDataGate | undefined;
  private transport: SheetsHttpTransport;
  private log: (message: string) => void;
  private sentRequests: GoogleRequest[];
  private sendSnapshotTaker: SendSnapshotTaker;
  constructor(props: NodeHostProps) {
    this.configs = props.configs;
    this.endpoints = props.endpoints ?? {};
    this.spreadsheetId = props.spreadsheetId;
    this.isDryRun = props.isDryRun;
    this.enteredDataGate = gateFor(props.enteredDataGate);
    this.transport = props.transport;
    this.log = props.log;
    this.sentRequests = [];
    this.sendSnapshotTaker = SendSnapshotTaker.init({
      spreadsheetId: props.spreadsheetId,
      transport: props.transport,
    });
  }
  static init(props: NodeHostProps): NodeHost {
    return new NodeHost(props);
  }
  get googleSheetsAPI(): GoogleSheetsAPI {
    return GoogleSheetsAPI.initHttp({
      spreadsheetId: this.spreadsheetId,
      transport: this.transport,
      reportRead: (read) => this.enteredDataGate?.integrateRead(read),
      gateFlush: (requests, send) => this._gateFlush(requests, send),
    });
  }
  // Every request the run produced, sent or withheld by the dry run.
  get summary(): UpdateRequestSummary {
    return new UpdateRequestSummary({ requests: this.sentRequests });
  }
  // Every flush's held items and their token; undefined when the host has no gate.
  get heldList(): HeldList | undefined {
    return this.enteredDataGate?.heldList;
  }
  ensureGlobals(): this {
    const globals = globalThis as Record<string, unknown>;
    globals.Logger = { log: this.log };
    installConfigs(this.configs);
    installEndpoints(this.endpoints);
    installRawSource(this.googleSheetsAPI);
    return this;
  }
  // What the send's held items changed, read just before each flush that held them, and the post-send fingerprints; undefined when nothing was held.
  sendSnapshot(): SendSnapshot | undefined {
    return this.sendSnapshotTaker.finished();
  }
  private _gateFlush(
    requests: GoogleRequest[],
    send: () => GoogleBatchUpdateResponse,
  ): GoogleBatchUpdateResponse {
    this.sentRequests.push(...requests);
    const gate = this.enteredDataGate;
    gate?.classify(requests);
    if (this.isDryRun) return {};
    if (gate === undefined) return send();
    // The gate calls this only once a token covers the flush, so a refused flush takes no snapshot.
    return gate.flush(() => {
      this.sendSnapshotTaker.capture(gate.snapshotPlan);
      return send();
    });
  }
}

function gateFor(
  props: EnteredDataGateProps | undefined,
): EnteredDataGate | undefined {
  if (props === undefined) return undefined;
  return EnteredDataGate.init(props);
}
