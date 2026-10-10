// `sheets-framework configs-diff`: summarises how the generated config files on disk differ from their HEAD versions. Read-only; no credential.
import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { relative } from "node:path";

import { Val } from "@byronbroughten/utils/val";

import {
  type ConfigFile,
  configFilePath,
  configFiles,
  hasPackageConfigs,
} from "./nodeHost.ts";
import type { SheetsConfig } from "./sheetsConfig.ts";

type ConfigEntry = Record<string, unknown>;
type TableEntry = ConfigEntry & { tableId: string };
type ColumnEntry = ConfigEntry & { columnId: string };

export interface ConfigSnapshot {
  tableConfigs: Record<string, TableEntry>;
  columnConfigs: Record<string, Record<string, ColumnEntry>>;
  valueConfigs: Record<string, string[]>;
}

interface Section {
  header: string;
  changes: string[];
  lines: string[];
}

export function summarizeConfigsDiff(
  before: ConfigSnapshot | undefined,
  after: ConfigSnapshot,
): string[] {
  const sections = [
    ...diffTables(before ?? emptySnapshot(), after),
    ...diffValues(before?.valueConfigs ?? {}, after.valueConfigs),
  ];
  const count = sections.reduce(
    (total, { changes, lines }) =>
      total + (changes.length === 0 ? 0 : 1) + lines.length,
    0,
  );
  if (count === 0) return ["No config changes."];
  return [
    ...sections.flatMap(({ header, changes, lines }) => [
      changes.length === 0 ? `${header}:` : `${header}: ${changes.join("; ")}`,
      ...lines,
    ]),
    `${count} ${count === 1 ? "change" : "changes"}.`,
  ];
}

function emptySnapshot(): ConfigSnapshot {
  return { tableConfigs: {}, columnConfigs: {}, valueConfigs: {} };
}

function diffTables(before: ConfigSnapshot, after: ConfigSnapshot): Section[] {
  const afterKeyById = keyById(after.tableConfigs, "tableId");
  const beforeIds = new Set(
    Object.values(before.tableConfigs).map(({ tableId }) => tableId),
  );
  const matchedAndRemoved = Object.entries(before.tableConfigs).flatMap(
    ([beforeKey, beforeTable]) => {
      const afterKey = afterKeyById.get(beforeTable.tableId);
      const beforeCols = before.columnConfigs[beforeKey] ?? {};
      if (afterKey === undefined) {
        return [tableSection(beforeKey, ["removed"], diffCols(beforeCols, {}))];
      }
      const changes = entryChanges(
        { beforeKey, afterKey },
        beforeTable,
        Val.assert(after.tableConfigs[afterKey], "after table config"),
        "tableId",
      );
      const lines = diffCols(beforeCols, after.columnConfigs[afterKey] ?? {});
      return changes.length + lines.length === 0
        ? []
        : [tableSection(renameLabel(beforeKey, afterKey), changes, lines)];
    },
  );
  const added = Object.entries(after.tableConfigs)
    .filter(([, { tableId }]) => !beforeIds.has(tableId))
    .map(([key]) =>
      tableSection(
        key,
        ["added"],
        diffCols({}, after.columnConfigs[key] ?? {}),
      ),
    );
  return [...matchedAndRemoved, ...added];
}

function tableSection(
  label: string,
  changes: string[],
  lines: string[],
): Section {
  return { header: `Table ${label}`, changes, lines };
}

function diffCols(
  before: Record<string, ColumnEntry>,
  after: Record<string, ColumnEntry>,
): string[] {
  const afterKeyById = keyById(after, "columnId");
  const beforeIds = new Set(
    Object.values(before).map(({ columnId }) => columnId),
  );
  const matchedAndRemoved = Object.entries(before).flatMap(
    ([beforeKey, beforeCol]) => {
      const afterKey = afterKeyById.get(beforeCol.columnId);
      if (afterKey === undefined) return [colLine(beforeKey, ["removed"])];
      const changes = entryChanges(
        { beforeKey, afterKey },
        beforeCol,
        Val.assert(after[afterKey], "after column config"),
        "columnId",
      );
      return changes.length === 0
        ? []
        : [colLine(renameLabel(beforeKey, afterKey), changes)];
    },
  );
  const added = Object.entries(after)
    .filter(([, { columnId }]) => !beforeIds.has(columnId))
    .map(([key]) => colLine(key, ["added"]));
  return [...matchedAndRemoved, ...added];
}

function colLine(label: string, changes: string[]): string {
  return `  column ${label}: ${changes.join("; ")}`;
}

function keyById<EN extends ConfigEntry>(
  record: Record<string, EN>,
  idField: keyof EN,
): Map<unknown, string> {
  return new Map(
    Object.entries(record).map(([key, entry]) => [entry[idField], key]),
  );
}

function entryChanges(
  { beforeKey, afterKey }: { beforeKey: string; afterKey: string },
  before: ConfigEntry,
  after: ConfigEntry,
  idField: string,
): string[] {
  const renamed = beforeKey === afterKey ? [] : ["renamed"];
  const fields = [...new Set([...Object.keys(before), ...Object.keys(after)])];
  const fieldChanges = fields.flatMap((field) => {
    const old = JSON.stringify(before[field]);
    const next = JSON.stringify(after[field]);
    return field === idField || old === next
      ? []
      : [`${field} ${old} → ${next}`];
  });
  return [...renamed, ...fieldChanges];
}

function renameLabel(beforeKey: string, afterKey: string): string {
  return beforeKey === afterKey ? beforeKey : `${beforeKey} → ${afterKey}`;
}

function diffValues(
  before: Record<string, string[]>,
  after: Record<string, string[]>,
): Section[] {
  const matchedAndRemoved = Object.entries(before).flatMap(
    ([name, options]) => {
      const afterOptions = after[name];
      if (afterOptions === undefined) return [valueSection(name, ["removed"])];
      const changes = optionChanges(options, afterOptions);
      return changes.length === 0 ? [] : [valueSection(name, changes)];
    },
  );
  const added = Object.keys(after)
    .filter((name) => !(name in before))
    .map((name) => valueSection(name, ["added"]));
  return [...matchedAndRemoved, ...added];
}

function optionChanges(before: string[], after: string[]): string[] {
  const added = after.filter((option) => !before.includes(option));
  const removed = before.filter((option) => !after.includes(option));
  const changes = [
    ...(added.length === 0 ? [] : [`added ${quoteAll(added)}`]),
    ...(removed.length === 0 ? [] : [`removed ${quoteAll(removed)}`]),
  ];
  if (changes.length === 0 && before.join("\n") !== after.join("\n")) {
    return ["reordered"];
  }
  return changes;
}

function quoteAll(options: string[]): string {
  return options.map((option) => JSON.stringify(option)).join(", ");
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
  const after = readSnapshot((file) =>
    readFileSync(configFilePath(sheetsConfig.generatedDir, file), "utf8"),
  );
  const before = readSnapshot((file) =>
    readHeadFile(
      sheetsConfig.dir,
      configFilePath(sheetsConfig.generatedDir, file),
    ),
  );
  console.log(summarizeConfigsDiff(before, after).join("\n"));
}

function readSnapshot(
  readFile: (file: ConfigFile) => string | undefined,
): ConfigSnapshot {
  // A file missing at HEAD is a first generation of that file: all its entries count as added.
  const [tableConfigs, columnConfigs, valueConfigs] = configFiles.map(
    (file) => {
      const text = readFile(file);
      return text === undefined ? {} : parseConfigFile(text);
    },
  );
  return {
    tableConfigs: tableConfigs as ConfigSnapshot["tableConfigs"],
    columnConfigs: columnConfigs as ConfigSnapshot["columnConfigs"],
    valueConfigs: valueConfigs as ConfigSnapshot["valueConfigs"],
  };
}

// The generated files are one `make…Configs(<JSON>)` call; parsing the JSON avoids importing two versions of one module.
function parseConfigFile(text: string): unknown {
  const start = text.indexOf("(", text.indexOf("Configs = make"));
  const end = text.lastIndexOf(")");
  return JSON.parse(text.slice(start + 1, end));
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
