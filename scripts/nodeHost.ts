// Starts the framework's second host. See docs/how-it-runs/node-host.md.
import { spawnSync } from "node:child_process";
import { existsSync, readFileSync } from "node:fs";
import { homedir } from "node:os";
import { basename, join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

import { Val } from "@byronbroughten/utils/val";

import type { SheetsHttpRequest } from "../src/00_Source/GoogleSheets/GoogleSheetsAPI.ts";
import type {
  Configs,
  ConfigSetBase,
} from "../src/01_SpreadsheetSchema/configReaders/configRegister.ts";
import type { Endpoints } from "../src/06_API/Endpoints.ts";
import type { NodeHost } from "../src/nodeHost/NodeHost.ts";
import type { SheetsConfig } from "./sheetsConfig.ts";

const claspAuth = {
  user: "desktop-clasp-run",
  // clasp's own name for a `.clasprc.json`-format file's one login.
  agentUser: "default",
  fileEnvVar: "clasp_config_auth",
  tokenUrl: "https://oauth2.googleapis.com/token",
} as const;
const maxResponseBytes = 256 * 1024 * 1024;

// An app's configs register without the rest, so only these decide whether it has configs of its own.
const requiredConfigFiles = [
  "tableConfigs",
  "columnConfigs",
  "valueConfigs",
] as const;

export const configFiles = [
  ...requiredConfigFiles,
  "tabConfigs",
  "retiredIdPrefixes",
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
  label: string;
  client_id: string;
  client_secret: string;
  refresh_token: string;
}

const fetchSyncPath = fileURLToPath(new URL("./fetchSync.js", import.meta.url));

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

export async function startNodeHost({
  isDryRun,
  sheetsConfig,
  configs,
  endpoints,
}: {
  isDryRun: boolean;
  sheetsConfig: SheetsConfig;
  configs: Configs;
  endpoints?: Endpoints;
}): Promise<NodeHost> {
  const { NodeHost } = await import("../src/nodeHost/NodeHost.ts");
  const transport = SheetsTransport.init();
  return NodeHost.init({
    configs,
    endpoints,
    spreadsheetId: sheetsConfig.spreadsheetId,
    transport: (request) => transport.send(request),
    isDryRun,
    log: (message) => console.log(`  log: ${message}`),
  }).ensureGlobals();
}

// The generated files in the package's generatedDir, as the package's entry passes them to the framework.
export async function loadPackageConfigs({
  generatedDir,
}: SheetsConfig): Promise<Configs> {
  const configs: ConfigSetBase = {
    tableConfigs: await importConfig(generatedDir, "tableConfigs"),
    tabConfigs: await importConfigIfPresent(generatedDir, "tabConfigs"),
    retiredIdPrefixes: await importConfigIfPresent(
      generatedDir,
      "retiredIdPrefixes",
    ),
    columnConfigs: await importConfig(generatedDir, "columnConfigs"),
    valueConfigs: await importConfig(generatedDir, "valueConfigs"),
  };
  // The registered set may or may not carry the two optional files, and these are the generated files it registers.
  return configs as Configs;
}

async function importConfig<CF extends ConfigFile>(
  generatedDir: string,
  base: CF,
): Promise<ConfigSetBase[CF]> {
  const url = pathToFileURL(configFilePath(generatedDir, base));
  return (await import(url.href))[base];
}

// A file the package's last regen predates is absent, so its config reads as unset.
async function importConfigIfPresent<CF extends ConfigFile>(
  generatedDir: string,
  base: CF,
): Promise<ConfigSetBase[CF] | undefined> {
  const filePath = configFilePath(generatedDir, base);
  if (!existsSync(filePath)) return undefined;
  return (await import(pathToFileURL(filePath).href))[base];
}

export function hasPackageConfigs({ generatedDir }: SheetsConfig): boolean {
  return requiredConfigFiles.every((base) =>
    existsSync(configFilePath(generatedDir, base)),
  );
}

export function configFilePath(generatedDir: string, base: ConfigFile): string {
  return join(generatedDir, `${base}.ts`);
}

export async function loadPackageEndpoints({
  endpointModule,
}: SheetsConfig): Promise<Endpoints | undefined> {
  if (endpointModule === undefined) return undefined;
  return importConstNamedAfterFile<Endpoints>(
    endpointModule,
    "An endpoint module",
  );
}

// The one rule a chore file and an endpoint module share, so a typo names the file and the const it lacks.
export async function importConstNamedAfterFile<MX>(
  modulePath: string,
  fileKind: string,
): Promise<MX> {
  const name = basename(modulePath).replace(/\.ts$/, "");
  const value = (await import(pathToFileURL(modulePath).href))[name];
  if (!value) {
    throw new Error(
      `${modulePath} exports no "${name}". ${fileKind} exports one const named after the file.`,
    );
  }
  return value;
}

// The framework's own dev configs: enough to read any spreadsheet's config floor.
export async function loadFrameworkConfigs(): Promise<Configs> {
  return (await import("../dev/generated/appConfigs.ts")).appConfigs;
}
