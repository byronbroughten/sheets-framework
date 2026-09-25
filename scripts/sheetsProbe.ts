// `sheets-framework probe`: one read-only Sheets request, summarized; the full JSON goes to the package's .probe/last.json. See docs/how-it-runs.md, "Seeing the raw Sheets JSON".
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { join, relative } from "node:path";

import type { SheetsHttpRequest } from "../src/00_Source/GoogleSheets/GoogleSheetsAPI.ts";
import { SheetsTransport } from "./nodeHost.ts";
import type { SheetsConfig } from "./sheetsConfig.ts";

const sheetsApiBase = "https://sheets.googleapis.com/v4/spreadsheets";
const outputName = ".probe/last.json";
const printLimits = { lines: 60, listed: 50 } as const;

interface SheetsProbeProps {
  sheetsConfig: SheetsConfig;
  fields?: string;
  filter?: string;
  path?: string;
}

type ProbeOption = "fields" | "filter" | "path";

const probeFlags = new Map<string, ProbeOption>([
  ["--fields", "fields"],
  ["--filter", "filter"],
  ["--path", "path"],
]);

class SheetsProbe {
  readonly sheetsConfig: SheetsConfig;
  readonly fields: string | undefined;
  readonly filter: string | undefined;
  readonly path: string | undefined;
  readonly output: string;
  readonly outputShown: string;
  constructor({ sheetsConfig, fields, filter, path }: SheetsProbeProps) {
    this.sheetsConfig = sheetsConfig;
    this.fields = fields;
    this.filter = filter;
    this.path = path;
    this.output = join(sheetsConfig.dir, outputName);
    this.outputShown = outputShownOf(sheetsConfig);
  }
  static init(sheetsConfig: SheetsConfig, argv: string[]): SheetsProbe {
    const options: SheetsProbeProps = { sheetsConfig };
    for (let i = 0; i < argv.length; i += 2) {
      const flag = argv[i] ?? "";
      const option = probeFlags.get(flag);
      const value = argv[i + 1];
      if (option === undefined || value === undefined) {
        throw new Error(
          `Unexpected argument "${flag}".\n\n${usage(outputShownOf(sheetsConfig))}`,
        );
      }
      options[option] = value;
    }
    return new SheetsProbe(options);
  }
  run(): void {
    const isRequest = this.fields !== undefined || this.filter !== undefined;
    if (!isRequest && this.path === undefined) {
      console.log(usage(this.outputShown));
      return;
    }
    const response = isRequest ? this._fetch() : this._lastResponse();
    const value = valueAt(response, this.path);
    console.log(`\n${this.path ? `at ${this.path}` : "response"}:`);
    console.log(printed(value, this.outputShown));
  }
  _fetch(): unknown {
    const request = this._request();
    console.log(`probe: ${request.method} ${request.url}`);
    assertIsRead(request);
    const response = SheetsTransport.init().send(request);
    const json = JSON.stringify(response, null, 2);
    mkdirSync(join(this.sheetsConfig.dir, ".probe"), { recursive: true });
    writeFileSync(this.output, `${json}\n`);
    console.log(
      `saved: ${this.outputShown} (${json.split("\n").length} lines) — Read a line range of it when this summary is not enough.`,
    );
    return response;
  }
  // The only two requests this can build are reads; there is no way to name another verb.
  _request(): SheetsHttpRequest {
    const base = `${sheetsApiBase}/${this.sheetsConfig.spreadsheetId}`;
    const query = this.fields
      ? `?fields=${encodeURIComponent(this.fields)}`
      : "";
    if (this.filter === undefined) {
      return { method: "GET", url: `${base}${query}`, body: null };
    }
    return {
      method: "POST",
      url: `${base}:getByDataFilter${query}`,
      body: JSON.stringify(parsedFilter(this.filter)),
    };
  }
  _lastResponse(): unknown {
    try {
      return JSON.parse(readFileSync(this.output, "utf8"));
    } catch {
      throw new Error(
        `No saved response at ${this.outputShown}. Pass --fields or --filter to fetch one.`,
      );
    }
  }
}

// Shown from where npm was invoked, so dev's file reads as dev/.probe/last.json from the root.
function outputShownOf(sheetsConfig: SheetsConfig): string {
  return relative(
    process.env.INIT_CWD ?? process.cwd(),
    join(sheetsConfig.dir, outputName),
  );
}

function usage(outputShown: string): string {
  return `Usage:
  npm run <app|dev>:probe -- --fields '<mask>' [--path <path>]         GET the spreadsheet with a fields mask
  npm run <app|dev>:probe -- --filter '<getByDataFilter body JSON>' [--fields '<mask>'] [--path <path>]
  npm run <app|dev>:probe -- --path <path>                             re-read ${outputShown}, no request

A path is dot-separated: an index, a key, or key=value to pick an array
element by that key or by properties.<key> — e.g. sheets.title=Occupancy.protectedRanges.
Only a summary is printed; the full response is saved to ${outputShown}.`;
}

// A backstop for future edits to _request: the transport itself has no dry-run gate.
function assertIsRead({ method, url }: SheetsHttpRequest): void {
  const isGet =
    method === "GET" && !/:\w+(\?|$)/.test(url.slice(url.lastIndexOf("/") + 1));
  const isFilterRead = method === "POST" && /:getByDataFilter(\?|$)/.test(url);
  if (!isGet && !isFilterRead) {
    throw new Error(`The probe only reads; refusing ${method} ${url}.`);
  }
}

function parsedFilter(filter: string): unknown {
  let body;
  try {
    body = JSON.parse(filter);
  } catch (error) {
    throw new Error(`--filter is not JSON: ${error.message}`);
  }
  if (!Array.isArray(body?.dataFilters)) {
    throw new Error('--filter needs a "dataFilters" array.');
  }
  return body;
}

function valueAt(root: unknown, path: string | undefined): unknown {
  if (!path) return root;
  let value = root;
  path.split(".").forEach((segment) => {
    value = childAt(value, segment);
    if (value === undefined) {
      throw new Error(`Nothing at "${segment}" in path "${path}".`);
    }
  });
  return value;
}

function childAt(value: unknown, segment: string): unknown {
  if (value === null || typeof value !== "object") return undefined;
  const match = /^([^=]+)=(.*)$/.exec(segment);
  if (!match || !Array.isArray(value)) return fieldOf(value, segment);
  const [, key = "", wanted] = match;
  return value.find(
    (each) =>
      String(
        fieldOf(each, key) ?? fieldOf(fieldOf(each, "properties"), key),
      ) === wanted,
  );
}

// Small subtrees print whole; anything longer prints as a shape summary.
function printed(value: unknown, outputShown: string): string {
  const json = JSON.stringify(value, null, 2) ?? "undefined";
  const lines = json.split("\n");
  if (lines.length <= printLimits.lines) return json;
  return `${shapeLines(value).join("\n")}\n\n(${lines.length} lines as JSON — narrow with --path, or Read ${outputShown} with offset/limit.)`;
}

function shapeLines(value: unknown): string[] {
  if (Array.isArray(value)) return arrayLines(value);
  if (value === null || typeof value !== "object") return [shapeOf(value)];
  return Object.entries(value).flatMap(([key, child]) => [
    `  ${key}: ${shapeOf(child)}`,
    ...(Array.isArray(child) ? arrayLines(child).slice(1) : []),
  ]);
}

function arrayLines(array: unknown[]): string[] {
  const lines = [shapeOf(array)];
  array.slice(0, printLimits.listed).forEach((each, index) => {
    const label = labelOf(each);
    const counts = countsOf(each);
    if (label || counts) lines.push(`    ${index}: ${label}${counts}`);
  });
  if (array.length > printLimits.listed) {
    lines.push(`    … ${array.length - printLimits.listed} more`);
  }
  return lines;
}

function labelOf(value: unknown): string {
  const source = fieldOf(value, "properties") ?? value;
  if (source === null || typeof source !== "object") return "";
  const parts = ["sheetId", "title", "tableId", "name"]
    .filter((key) => fieldOf(source, key) !== undefined)
    .map((key) => `${key}=${JSON.stringify(fieldOf(source, key))}`);
  return parts.join(" ");
}

function countsOf(value: unknown): string {
  if (value === null || typeof value !== "object" || Array.isArray(value)) {
    return "";
  }
  const counts = Object.entries(value)
    .filter(([, child]) => Array.isArray(child))
    .map(([key, child]) => `${key}[${child.length}]`);
  return counts.length > 0 ? `  ${counts.join(" ")}` : "";
}

function shapeOf(value: unknown): string {
  if (Array.isArray(value)) {
    const keys = new Set(
      value.flatMap((each) =>
        each && typeof each === "object" && !Array.isArray(each)
          ? Object.keys(each)
          : [],
      ),
    );
    return `array[${value.length}]${keys.size > 0 ? ` of {${[...keys].join(", ")}}` : ""}`;
  }
  if (value !== null && typeof value === "object") {
    return `object {${Object.keys(value).join(", ")}}`;
  }
  const json = JSON.stringify(value) ?? "undefined";
  return json.length > 60 ? `${json.slice(0, 57)}…` : json;
}

// Optional chaining's semantics on a parsed JSON value of unknown shape.
function fieldOf(value: unknown, key: string): unknown {
  return (value as Record<string, unknown> | null | undefined)?.[key];
}

export function runProbe(sheetsConfig: SheetsConfig, argv: string[]): void {
  SheetsProbe.init(sheetsConfig, argv).run();
}
