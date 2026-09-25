// Starts the framework's second host. See docs/how-it-runs.md.
import { spawnSync } from "node:child_process";
import { existsSync, readFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

import type { SheetsHttpRequest } from "../src/00_Source/GoogleSheets/GoogleSheetsAPI.ts";
import type { Configs } from "../src/01_SpreadsheetSchema/configRegister.ts";
import type { NodeHost } from "../src/nodeHost/NodeHost.ts";
import { Val } from "../src/utils/Val.ts";
import type { SheetsConfig } from "./sheetsConfig.ts";

const claspAuth = {
  user: "desktop-clasp-run",
  tokenUrl: "https://oauth2.googleapis.com/token",
} as const;
const maxResponseBytes = 256 * 1024 * 1024;

export const configFiles = [
  "spreadsheetConfig",
  "sheetConfigs",
  "columnConfigs",
  "valueConfigs",
] as const;

export type ConfigFile = (typeof configFiles)[number];

interface FetchRequest {
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
  client_id: string;
  client_secret: string;
  refresh_token: string;
}

const path = {
  fetchSync: fileURLToPath(new URL("./fetchSync.js", import.meta.url)),
  claspRc: fileURLToPath(new URL(".clasprc.json", `file://${homedir()}/`)),
};

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
        `Could not refresh the "${claspAuth.user}" token (HTTP ${status}). ` +
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

function claspCredential(): ClaspCredential {
  let tokens;
  try {
    ({ tokens } = JSON.parse(readFileSync(path.claspRc, "utf8")));
  } catch (error) {
    throw new Error(
      `Could not read ${path.claspRc}: ${error.message}. ` +
        "Run sheets-framework setup-auth.",
    );
  }
  const credential = tokens?.[claspAuth.user];
  if (!credential?.refresh_token) {
    throw new Error(
      `No "${claspAuth.user}" credential in ${path.claspRc}. ` +
        "Run sheets-framework setup-auth.",
    );
  }
  return {
    client_id: Val.validate.string(credential.client_id),
    client_secret: Val.validate.string(credential.client_secret),
    refresh_token: Val.validate.string(credential.refresh_token),
  };
}

function fetchSync(request: FetchRequest): FetchResponse {
  const { status, stdout, stderr, error } = spawnSync(
    process.execPath,
    [path.fetchSync],
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

export async function startNodeHost({
  isDryRun,
  sheetsConfig,
  configs,
}: {
  isDryRun: boolean;
  sheetsConfig: SheetsConfig;
  configs: Configs;
}): Promise<NodeHost> {
  const { NodeHost } = await import("../src/nodeHost/NodeHost.ts");
  const transport = SheetsTransport.init();
  return NodeHost.init({
    configs,
    spreadsheetId: sheetsConfig.spreadsheetId,
    transport: (request) => transport.send(request),
    isDryRun,
    log: (message) => console.log(`  log: ${message}`),
  }).ensureGlobals();
}

// The four files in the package's generatedDir, as the package's entry passes them to the framework.
export async function loadPackageConfigs({
  generatedDir,
}: SheetsConfig): Promise<Configs> {
  return {
    spreadsheetConfig: await importConfig(generatedDir, "spreadsheetConfig"),
    sheetConfigs: await importConfig(generatedDir, "sheetConfigs"),
    columnConfigs: await importConfig(generatedDir, "columnConfigs"),
    valueConfigs: await importConfig(generatedDir, "valueConfigs"),
  };
}

async function importConfig<CF extends ConfigFile>(
  generatedDir: string,
  base: CF,
): Promise<Configs[CF]> {
  const url = pathToFileURL(configFilePath(generatedDir, base));
  return (await import(url.href))[base];
}

export function hasPackageConfigs({ generatedDir }: SheetsConfig): boolean {
  return configFiles.every((base) =>
    existsSync(configFilePath(generatedDir, base)),
  );
}

export function configFilePath(generatedDir: string, base: ConfigFile): string {
  return join(generatedDir, `${base}.ts`);
}

// The framework's own dev configs: enough to read any spreadsheet's config floor.
export async function loadFrameworkConfigs(): Promise<Configs> {
  return (await import("../dev/devConfigs.ts")).devConfigs;
}
