import { spawnSync } from "node:child_process";
import { existsSync, readFileSync } from "node:fs";
import { relative } from "node:path";

import { Obj } from "@byronbroughten/utils/obj";
import { Val } from "@byronbroughten/utils/val";

import type {
  FixedListMember,
  ValueConfigBase,
} from "../../src/01_SpreadsheetSchema/configReaders/makeConfigs.ts";
import type {
  ConfigSnapshot,
  IdentifiedEntry,
  IdField,
  RuleGroups,
  RuleSnapshot,
  TabEntry,
} from "../configsDiff.ts";
import type { ConfigFile } from "../packageLoader.ts";
import { valueKindLabels } from "./snapshotDiff.ts";

type ReadFile = (file: ConfigFile) => string | undefined;

export const configFiles = {
  readSnapshot(readFile: ReadFile): ConfigSnapshot {
    return {
      tableConfigs: mapValues(
        readConfigFile(readFile, "tableConfigs"),
        (entry) => validateEntry(entry, "tableId"),
      ),
      tabConfigs: mapValues(
        readConfigFile(readFile, "tabConfigs"),
        validateTabEntry,
      ),
      retiredIdPrefixes: readRetiredIdPrefixes(readFile),
      columnConfigs: mapValues(
        readConfigFile(readFile, "columnConfigs"),
        (columns) =>
          mapValues(
            validateRecord(columns, "a Table's columnConfigs"),
            (entry) => validateEntry(entry, "columnId"),
          ),
      ),
      valueConfigs: mapValues(
        readConfigFile(readFile, "valueConfigs"),
        validateValueConfig,
      ),
      conditionalFormatConfigs: readRuleSnapshot(readFile),
    };
  },
  // A package whose last regen predates a generated file has none on disk yet.
  readFileIfPresent(path: string): string | undefined {
    if (!existsSync(path)) return undefined;
    return readFileSync(path, "utf8");
  },
  readHeadFile(cwd: string, path: string): string | undefined {
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
  },
};

function mapValues<FR, TO>(
  record: Record<string, FR>,
  transform: (value: FR) => TO,
): Record<string, TO> {
  return Object.fromEntries(
    Object.entries(record).map(([key, value]) => [key, transform(value)]),
  );
}

function readConfigFile(
  readFile: ReadFile,
  file: ConfigFile,
): Record<string, unknown> {
  const text = readFile(file);
  // A file missing at HEAD is a first generation of that file: all its entries count as added.
  if (text === undefined) return {};
  return validateRecord(parseConfigFile(text), file);
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

function readRetiredIdPrefixes(readFile: ReadFile): string[] {
  const text = readFile("retiredIdPrefixes");
  if (text === undefined) return [];
  const value = parseConfigFile(text);
  if (Array.isArray(value)) {
    return value.map((prefix) => Val.validate.string(prefix));
  }
  throw new Error("configs:diff: retiredIdPrefixes is not a JSON array.");
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

function readRuleSnapshot(readFile: ReadFile): RuleSnapshot {
  const file = readConfigFile(readFile, "conditionalFormatConfigs");
  return {
    tables: readRuleGroups(file.tables, "conditionalFormatConfigs.tables"),
    tabs: readRuleGroups(file.tabs, "conditionalFormatConfigs.tabs"),
  };
}

function readRuleGroups(value: unknown, what: string): RuleGroups {
  if (value === undefined) return {};
  return mapValues(validateRecord(value, what), (group) =>
    mapValues(validateRecord(group, `a group of ${what}`), (rule) =>
      validateRecord(rule, "a conditional format rule"),
    ),
  );
}
