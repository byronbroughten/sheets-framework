// `sheets-framework configs-diff`: summarises how the generated config files on disk differ from their HEAD versions. Read-only; no credential.
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { relative } from "node:path";

import { Obj } from "@byronbroughten/utils/obj";
import { Val } from "@byronbroughten/utils/val";

import {
  type ConfigFile,
  configFilePath,
  hasPackageConfigs,
} from "./nodeHost.ts";
import type { SheetsConfig } from "./sheetsConfig.ts";

type ConfigEntry = Record<string, unknown>;
type IdField = "tableId" | "columnId";
type IdentifiedEntry<FN extends IdField> = ConfigEntry & Record<FN, string>;

export interface ConfigSnapshot {
  tableConfigs: Record<string, IdentifiedEntry<"tableId">>;
  columnConfigs: Record<string, Record<string, IdentifiedEntry<"columnId">>>;
  valueConfigs: Record<string, string[]>;
}

interface BeforeAfter<SV> {
  before: SV;
  after: SV;
}

export interface SnapshotPair {
  before: ConfigSnapshot | undefined;
  after: ConfigSnapshot;
}

interface Matched<EN extends object> extends BeforeAfter<EN> {
  beforeKey: string;
  afterKey: string;
}

type Pairing<EN extends object> =
  | ({ kind: "matched" } & Matched<EN>)
  | { kind: "removed"; key: string; entry: EN }
  | { kind: "added"; key: string; entry: EN };

interface Section {
  header: string;
  changes: string[];
  lines: string[];
}

export function summarizeConfigsDiff({
  before = emptySnapshot(),
  after,
}: SnapshotPair): string[] {
  const sections = [
    ...diffTables({ before, after }),
    ...diffValues({ before: before.valueConfigs, after: after.valueConfigs }),
  ];
  const headerCount = sections.filter(
    ({ changes }) => changes.length > 0,
  ).length;
  const lineCount = sections.reduce(
    (total, { lines }) => total + lines.length,
    0,
  );
  const count = headerCount + lineCount;
  if (count === 0) return ["No config changes."];
  return [
    ...sections.flatMap((section) => [
      sectionHeader(section),
      ...section.lines,
    ]),
    `${count} ${count === 1 ? "change" : "changes"}.`,
  ];
}

function sectionHeader({ header, changes }: Section): string {
  if (changes.length === 0) return `${header}:`;
  return `${header}: ${changes.join("; ")}`;
}

function emptySnapshot(): ConfigSnapshot {
  return { tableConfigs: {}, columnConfigs: {}, valueConfigs: {} };
}

function diffTables({ before, after }: BeforeAfter<ConfigSnapshot>): Section[] {
  const pairings = pairEntries(
    { before: before.tableConfigs, after: after.tableConfigs },
    (_key, { tableId }) => tableId,
  );
  return pairings.flatMap((pairing) => {
    if (pairing.kind === "removed") {
      const columns = before.columnConfigs[pairing.key] ?? {};
      const lines = diffCols({ before: columns, after: {} });
      return [tableSection(pairing.key, ["removed"], lines)];
    }
    if (pairing.kind === "added") {
      const columns = after.columnConfigs[pairing.key] ?? {};
      const lines = diffCols({ before: {}, after: columns });
      return [tableSection(pairing.key, ["added"], lines)];
    }
    const changes = entryChanges(pairing, "tableId");
    const lines = diffCols({
      before: before.columnConfigs[pairing.beforeKey] ?? {},
      after: after.columnConfigs[pairing.afterKey] ?? {},
    });
    if (changes.length + lines.length === 0) return [];
    return [tableSection(renameLabel(pairing), changes, lines)];
  });
}

function tableSection(
  label: string,
  changes: string[],
  lines: string[],
): Section {
  return { header: `Table ${label}`, changes, lines };
}

function diffCols(
  columns: BeforeAfter<Record<string, IdentifiedEntry<"columnId">>>,
): string[] {
  const pairings = pairEntries(columns, (_key, { columnId }) => columnId);
  return pairings.flatMap((pairing) => {
    if (pairing.kind !== "matched") {
      return [colLine(pairing.key, [pairing.kind])];
    }
    const changes = entryChanges(pairing, "columnId");
    if (changes.length === 0) return [];
    return [colLine(renameLabel(pairing), changes)];
  });
}

function colLine(label: string, changes: string[]): string {
  return `  column ${label}: ${changes.join("; ")}`;
}

// Removed and matched entries keep the before order; added entries follow in the after order.
function pairEntries<EN extends object>(
  { before, after }: BeforeAfter<Record<string, EN>>,
  identity: (key: string, entry: EN) => unknown,
): Pairing<EN>[] {
  const afterKeyById = new Map(
    Object.entries(after).map(([key, entry]) => [identity(key, entry), key]),
  );
  const beforeIds = new Set(
    Object.entries(before).map(([key, entry]) => identity(key, entry)),
  );
  const matchedAndRemoved = Object.entries(before).map(
    ([beforeKey, entry]): Pairing<EN> => {
      const afterKey = afterKeyById.get(identity(beforeKey, entry));
      if (afterKey === undefined) {
        return { kind: "removed", key: beforeKey, entry };
      }
      return {
        kind: "matched",
        beforeKey,
        afterKey,
        before: entry,
        after: Val.assert(after[afterKey], "after config entry"),
      };
    },
  );
  const added = Object.entries(after)
    .filter(([key, entry]) => !beforeIds.has(identity(key, entry)))
    .map(([key, entry]): Pairing<EN> => ({ kind: "added", key, entry }));
  return [...matchedAndRemoved, ...added];
}

function entryChanges(
  { beforeKey, afterKey, before, after }: Matched<ConfigEntry>,
  idField: IdField,
): string[] {
  const renamed = beforeKey === afterKey ? [] : ["renamed"];
  const fields = [...new Set([...Object.keys(before), ...Object.keys(after)])];
  const fieldChanges = fields.flatMap((field) => {
    const old = JSON.stringify(before[field]);
    const next = JSON.stringify(after[field]);
    if (field === idField || old === next) return [];
    return [`${field} ${old} → ${next}`];
  });
  return [...renamed, ...fieldChanges];
}

function renameLabel({ beforeKey, afterKey }: Matched<object>): string {
  return beforeKey === afterKey ? beforeKey : `${beforeKey} → ${afterKey}`;
}

function diffValues(values: BeforeAfter<Record<string, string[]>>): Section[] {
  return pairEntries(values, (name) => name).flatMap((pairing) => {
    if (pairing.kind !== "matched") {
      return [valueSection(pairing.key, [pairing.kind])];
    }
    const changes = optionChanges(pairing);
    if (changes.length === 0) return [];
    return [valueSection(pairing.beforeKey, changes)];
  });
}

function optionChanges({ before, after }: BeforeAfter<string[]>): string[] {
  const added = after.filter((option) => !before.includes(option));
  const removed = before.filter((option) => !after.includes(option));
  const changes = [
    ...optionListChange("added", added),
    ...optionListChange("removed", removed),
  ];
  if (changes.length === 0 && before.join("\n") !== after.join("\n")) {
    return ["reordered"];
  }
  return changes;
}

function optionListChange(verb: string, options: string[]): string[] {
  if (options.length === 0) return [];
  const quoted = options.map((option) => JSON.stringify(option));
  return [`${verb} ${quoted.join(", ")}`];
}

function valueSection(name: string, changes: string[]): Section {
  return { header: `Value ${name}`, changes, lines: [] };
}

export function runConfigsDiff(sheetsConfig: SheetsConfig): void {
  if (!hasPackageConfigs(sheetsConfig)) {
    throw new Error(
      `configs:diff: no generated config files in ${sheetsConfig.generatedDir}; run gen:configs first.`,
    );
  }
  console.log(readConfigsDiff(sheetsConfig).join("\n"));
}

export function readConfigsDiff(sheetsConfig: SheetsConfig): string[] {
  const { dir, generatedDir } = sheetsConfig;
  return summarizeConfigsDiff({
    before: readSnapshot((file) =>
      readHeadFile(dir, configFilePath(generatedDir, file)),
    ),
    after: readSnapshot((file) =>
      readFileSync(configFilePath(generatedDir, file), "utf8"),
    ),
  });
}

function readSnapshot(
  readFile: (file: ConfigFile) => string | undefined,
): ConfigSnapshot {
  return {
    tableConfigs: mapValues(readConfigFile(readFile, "tableConfigs"), (entry) =>
      validateEntry(entry, "tableId"),
    ),
    columnConfigs: mapValues(
      readConfigFile(readFile, "columnConfigs"),
      (columns) =>
        mapValues(validateRecord(columns, "a Table's columnConfigs"), (entry) =>
          validateEntry(entry, "columnId"),
        ),
    ),
    valueConfigs: mapValues(
      readConfigFile(readFile, "valueConfigs"),
      validateOptions,
    ),
  };
}

function readConfigFile(
  readFile: (file: ConfigFile) => string | undefined,
  file: ConfigFile,
): Record<string, unknown> {
  const text = readFile(file);
  // A file missing at HEAD is a first generation of that file: all its entries count as added.
  if (text === undefined) return {};
  return validateRecord(parseConfigFile(text), file);
}

// The generated files are one `make…Configs(<JSON>)` call; parsing the JSON avoids importing two versions of one module.
function parseConfigFile(text: string): unknown {
  const start = text.indexOf("(", text.indexOf("Configs = make"));
  const end = text.lastIndexOf(")");
  return JSON.parse(text.slice(start + 1, end));
}

function validateRecord(value: unknown, what: string): Record<string, unknown> {
  if (Obj.isObjToRecord(value) && !Array.isArray(value)) return value;
  throw new Error(`configs:diff: ${what} is not a JSON object.`);
}

function validateEntry<FN extends IdField>(
  value: unknown,
  idField: FN,
): IdentifiedEntry<FN> {
  const entry = validateRecord(value, `a ${idField} entry`);
  Val.validate.string(entry[idField]);
  return entry as IdentifiedEntry<FN>;
}

function validateOptions(value: unknown): string[] {
  if (!Array.isArray(value)) {
    throw new Error("configs:diff: a valueConfigs entry is not an array.");
  }
  return value.map((option) => Val.validate.string(option));
}

function mapValues<FR, TO>(
  record: Record<string, FR>,
  transform: (value: FR) => TO,
): Record<string, TO> {
  return Object.fromEntries(
    Object.entries(record).map(([key, value]) => [key, transform(value)]),
  );
}

function readHeadFile(cwd: string, path: string): string | undefined {
  const { status, stdout } = spawnSync(
    "git",
    ["show", `HEAD:./${relative(cwd, path)}`],
    {
      cwd,
      encoding: "utf8",
      maxBuffer: 64 * 1024 * 1024,
    },
  );
  return status === 0 ? stdout : undefined;
}
