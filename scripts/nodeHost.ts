// Starts the framework's second host. See docs/how-it-runs/node-host.md.
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

import { Val } from "@byronbroughten/utils/val";

import type {
  SheetsHttpRequest,
  SheetsHttpTransport,
} from "../src/00_Source/GoogleSheets/GoogleSheetsAPI.ts";
import type { Configs } from "../src/01_SpreadsheetSchema/configReaders/configRegister.ts";
import type { Endpoints } from "../src/06_API/Endpoints.ts";
import type { EnteredDataGateProps } from "../src/nodeHost/EnteredDataGate.ts";
import type { NodeHost } from "../src/nodeHost/NodeHost.ts";
import { loadDropdownStyleReset } from "./packageLoader.ts";
import type { SheetsConfig } from "./sheetsConfig.ts";

const claspAuth = {
  user: "desktop-clasp-run",
  // clasp's own name for a `.clasprc.json`-format file's one login.
  agentUser: "default",
  fileEnvVar: "clasp_config_auth",
  tokenUrl: "https://oauth2.googleapis.com/token",
} as const;
const maxResponseBytes = 256 * 1024 * 1024;

export interface FetchRequest {
  url: string;
  method: string;
  headers: Record<string, string>;
  body: string | null;
}

interface FetchResponse {
  status: number;
  body: string;
}

interface ClaspCredential {
  label: string;
  client_id: string;
  client_secret: string;
  refresh_token: string;
}

const fetchSyncPath = fileURLToPath(new URL("./fetchSync.ts", import.meta.url));

// Node's built-in env proxy, so fetch reaches Google through the Bash sandbox's proxy with no dependency.
const fetchSyncFlags = ["--use-env-proxy", "--disable-warning=UNDICI-EHPA"];

export class SheetsTransport {
  accessToken: string | null = null;
  static init(): SheetsTransport {
    return new SheetsTransport();
  }
  send(request: SheetsHttpRequest): unknown {
    const { status, body } = fetchSync({
      url: request.url,
      method: request.method,
      headers: {
        Authorization: `Bearer ${this._ensureAccessToken()}`,
        "Content-Type": "application/json",
      },
      body: request.body,
    });
    if (status < 200 || status >= 300) {
      throw new Error(
        `Sheets API ${status} for ${request.method} ${request.url}\n${body}`,
      );
    }
    return JSON.parse(body);
  }
  _ensureAccessToken(): string {
    if (this.accessToken) return this.accessToken;
    const credential = claspCredential();
    const { status, body } = fetchSync({
      url: claspAuth.tokenUrl,
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: credential.client_id,
        client_secret: credential.client_secret,
        refresh_token: credential.refresh_token,
        grant_type: "refresh_token",
      }).toString(),
    });
    const accessToken = status === 200 ? accessTokenOf(body) : null;
    if (!accessToken) {
      throw new Error(
        `Could not refresh the token of the ${credential.label} (HTTP ${status}). ` +
          `Run sheets-framework setup-auth to re-mint it.\n${body}`,
      );
    }
    this.accessToken = accessToken;
    return accessToken;
  }
}

// A failing token endpoint need not answer in JSON, so a parse error is one too.
function accessTokenOf(body: string): string | null {
  try {
    const { access_token: accessToken } = JSON.parse(body);
    return Val.is.string(accessToken) ? accessToken : null;
  } catch {
    return null;
  }
}

// The agent credential file when `clasp_config_auth` names one, else the developer's own clasp login.
export function claspCredential(
  env: NodeJS.ProcessEnv = process.env,
  homeDir: string = homedir(),
): ClaspCredential {
  const agentFile = env[claspAuth.fileEnvVar];
  const claspRc = agentFile || join(homeDir, ".clasprc.json");
  const user = agentFile ? claspAuth.agentUser : claspAuth.user;
  let tokens;
  try {
    ({ tokens } = JSON.parse(readFileSync(claspRc, "utf8")));
  } catch (error) {
    throw new Error(
      `Could not read ${claspRc}: ${error.message}. ` +
        "Run sheets-framework setup-auth.",
    );
  }
  const credential = tokens?.[user];
  if (!credential?.refresh_token) {
    throw new Error(
      `No "${user}" credential in ${claspRc}. ` +
        "Run sheets-framework setup-auth.",
    );
  }
  return {
    label: `"${user}" credential in ${claspRc}`,
    client_id: Val.validate.string(credential.client_id),
    client_secret: Val.validate.string(credential.client_secret),
    refresh_token: Val.validate.string(credential.refresh_token),
  };
}

function fetchSync(request: FetchRequest): FetchResponse {
  const { status, stdout, stderr, error } = spawnSync(
    process.execPath,
    [...fetchSyncFlags, fetchSyncPath],
    {
      input: JSON.stringify(request),
      encoding: "utf8",
      maxBuffer: maxResponseBytes,
    },
  );
  if (error) throw error;
  if (status !== 0) {
    throw new Error(
      `The request subprocess exited with ${status}.${stderr.trim() ? ` ${stderr.trim()}` : ""}`,
    );
  }
  const response = JSON.parse(stdout);
  return {
    status: Val.validate.number(response.status),
    body: Val.validate.string(response.body),
  };
}

// A test's stand-ins for Google, for the package's configs, which a program installs once, and for the clock.
export interface HostStandIn {
  transport: SheetsHttpTransport;
  configs: Configs;
  now?: () => Date;
}

// The gate's terms but what the Node half supplies: its hash and the app's dropdownStyleReset.
type EnteredDataGateTerms = Omit<
  EnteredDataGateProps,
  "hash" | "dropdownStyleReset"
>;

export async function startNodeHost({
  isDryRun,
  sheetsConfig,
  configs,
  endpoints,
  enteredDataGate,
  transport = liveTransport(),
}: {
  isDryRun: boolean;
  sheetsConfig: SheetsConfig;
  configs: Configs;
  endpoints?: Endpoints;
  enteredDataGate?: EnteredDataGateTerms;
  // A test's stand-in for Google; the live one otherwise.
  transport?: SheetsHttpTransport;
}): Promise<NodeHost> {
  const { NodeHost } = await import("../src/nodeHost/NodeHost.ts");
  return NodeHost.init({
    configs,
    endpoints,
    spreadsheetId: sheetsConfig.spreadsheetId,
    transport,
    isDryRun,
    log: (message) => console.log(`  log: ${message}`),
    enteredDataGate: await gateProps(sheetsConfig, enteredDataGate),
  }).ensureGlobals();
}

export function liveTransport(): SheetsHttpTransport {
  const transport = SheetsTransport.init();
  return (request) => transport.send(request);
}

async function gateProps(
  sheetsConfig: SheetsConfig,
  terms: EnteredDataGateTerms | undefined,
): Promise<EnteredDataGateProps | undefined> {
  if (terms === undefined) return undefined;
  return {
    ...terms,
    dropdownStyleReset: await loadDropdownStyleReset(sheetsConfig),
    hash: (input) => createHash("sha256").update(input).digest("hex"),
  };
}
