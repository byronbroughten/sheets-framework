// `sheets-framework configs-diff`: summarises how the generated config files on disk differ from their HEAD versions. Read-only; no credential.
import { spawnSync } from "node:child_process";
import { existsSync, readFileSync } from "node:fs";
import { relative } from "node:path";

import { Obj } from "@byronbroughten/utils/obj";
import { Val } from "@byronbroughten/utils/val";

import type {
  FixedList,
  FixedListMember,
  ValueConfigBase,
} from "../src/01_SpreadsheetSchema/configReaders/makeConfigs.ts";
import {
  type ConfigFile,
  configFilePath,
  hasPackageConfigs,
} from "./nodeHost.ts";
import type { SheetsConfig } from "./sheetsConfig.ts";

type ConfigEntry = Record<string, unknown>;
type IdField = "tableId" | "columnId";
type IdentifiedEntry<FN extends IdField> = ConfigEntry & Record<FN, string>;
type TabEntry = ConfigEntry & { sheetGid: number };

export interface ConfigSnapshot {
  tableConfigs: Record<string, IdentifiedEntry<"tableId">>;
  tabConfigs: Record<string, TabEntry>;
  retiredIdPrefixes: string[];
  columnConfigs: Record<string, Record<string, IdentifiedEntry<"columnId">>>;
  valueConfigs: Record<string, ValueConfigBase>;
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
    ...diffTabs({ before: before.tabConfigs, after: after.tabConfigs }),
    ...diffRetiredIdPrefixes({
      before: before.retiredIdPrefixes,
      after: after.retiredIdPrefixes,
    }),
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
  return {
    tableConfigs: {},
    tabConfigs: {},
    retiredIdPrefixes: [],
    columnConfigs: {},
    valueConfigs: {},
  };
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

function diffTabs(tabs: BeforeAfter<Record<string, TabEntry>>): Section[] {
  const pairings = pairEntries(tabs, (_key, { sheetGid }) => sheetGid);
  return pairings.flatMap((pairing) => {
    if (pairing.kind !== "matched") {
      return [tabSection(pairing.key, [pairing.kind])];
    }
    const changes = entryChanges(pairing, "sheetGid");
    if (changes.length === 0) return [];
    return [tabSection(renameLabel(pairing), changes)];
  });
}

function tabSection(label: string, changes: string[]): Section {
  return { header: `Tab ${label}`, changes, lines: [] };
}

function diffRetiredIdPrefixes(prefixes: BeforeAfter<string[]>): Section[] {
  const changes = memberChanges(prefixes);
  if (changes.length === 0) return [];
  return [{ header: "Retired ID prefixes", changes, lines: [] }];
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
  idField: IdField | "sheetGid",
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

function diffValues(
  values: BeforeAfter<Record<string, ValueConfigBase>>,
): Section[] {
  return pairEntries(values, (name) => name).flatMap((pairing) => {
    if (pairing.kind !== "matched") {
      return [valueSection(pairing.key, [pairing.kind])];
    }
    const changes = valueConfigChanges(pairing);
    if (changes.length === 0) return [];
    return [valueSection(pairing.beforeKey, changes)];
  });
}

function valueConfigChanges({
  before,
  after,
}: BeforeAfter<ValueConfigBase>): string[] {
  if (isFixedList(before) && isFixedList(after)) {
    return memberChanges({ before, after });
  }
  const beforeKind = valueKindLabel(before);
  const afterKind = valueKindLabel(after);
  return beforeKind === afterKind ? [] : [`${beforeKind} → ${afterKind}`];
}

function isFixedList(config: ValueConfigBase): config is FixedList {
  return Array.isArray(config);
}

const valueKindLabels = {
  fixedList: "fixed list",
  liveReference: "live reference",
} as const;

function valueKindLabel(config: ValueConfigBase): string {
  if (isFixedList(config)) return valueKindLabels.fixedList;
  return valueKindLabels.liveReference;
}

function memberChanges({ before, after }: BeforeAfter<FixedList>): string[] {
  const added = after.filter((member) => !before.includes(member));
  const removed = before.filter((member) => !after.includes(member));
  const changes = [
    ...memberListChange("added", added),
    ...memberListChange("removed", removed),
  ];
  if (
    changes.length === 0 &&
    JSON.stringify(before) !== JSON.stringify(after)
  ) {
    return ["reordered"];
  }
  return changes;
}

function memberListChange(verb: string, members: FixedListMember[]): string[] {
  if (members.length === 0) return [];
  const quoted = members.map((member) => JSON.stringify(member));
  return [`${verb} ${quoted.join(", ")}`];
}

function valueSection(name: string, changes: string[]): Section {
  return { header: `Value ${name}`, changes, lines: [] };
}

export function runConfigsDiff(sheetsConfig: SheetsConfig): void {
  if (!hasPackageConfigs(sheetsConfig)) {
    throw new Error(
      `configs:diff: no generated config files in ${sheetsConfig.generatedDir}; run sync first.`,
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
      readFileIfPresent(configFilePath(generatedDir, file)),
    ),
  });
}

// A package whose last regen predates a generated file has none on disk yet.
function readFileIfPresent(path: string): string | undefined {
  if (!existsSync(path)) return undefined;
  return readFileSync(path, "utf8");
}

function readSnapshot(
  readFile: (file: ConfigFile) => string | undefined,
): ConfigSnapshot {
  return {
    tableConfigs: mapValues(readConfigFile(readFile, "tableConfigs"), (entry) =>
      validateEntry(entry, "tableId"),
    ),
    tabConfigs: mapValues(
      readConfigFile(readFile, "tabConfigs"),
      validateTabEntry,
    ),
    retiredIdPrefixes: readRetiredIdPrefixes(readFile),
    columnConfigs: mapValues(
      readConfigFile(readFile, "columnConfigs"),
      (columns) =>
        mapValues(validateRecord(columns, "a Table's columnConfigs"), (entry) =>
          validateEntry(entry, "columnId"),
        ),
    ),
    valueConfigs: mapValues(
      readConfigFile(readFile, "valueConfigs"),
      validateValueConfig,
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

function readRetiredIdPrefixes(
  readFile: (file: ConfigFile) => string | undefined,
): string[] {
  const text = readFile("retiredIdPrefixes");
  if (text === undefined) return [];
  const value = parseConfigFile(text);
  if (Array.isArray(value)) {
    return value.map((prefix) => Val.validate.string(prefix));
  }
  throw new Error("configs:diff: retiredIdPrefixes is not a JSON array.");
}

// The generated files are one `make…(<JSON>)` call; parsing the JSON avoids importing two versions of one module.
function parseConfigFile(text: string): unknown {
  const start = text.indexOf("(", text.indexOf(" = make"));
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

function validateTabEntry(value: unknown): TabEntry {
  const entry = validateRecord(value, "a tabConfigs entry");
  Val.validate.number(entry.sheetGid);
  return entry as TabEntry;
}

function validateValueConfig(value: unknown): ValueConfigBase {
  if (Array.isArray(value)) return value.map(validateMember);
  const entry = validateRecord(value, "a valueConfigs entry");
  if (entry.kind === "liveReference") return { kind: "liveReference" };
  const { fixedList, liveReference } = valueKindLabels;
  throw new Error(
    `configs:diff: a valueConfigs entry is neither a ${fixedList} nor a ${liveReference}.`,
  );
}

function validateMember(value: unknown): FixedListMember {
  if (Val.is.string(value) || Val.is.number(value) || Val.is.boolean(value)) {
    return value;
  }
  throw new Error(
    "configs:diff: a fixed list member is not a string, number or boolean.",
  );
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
