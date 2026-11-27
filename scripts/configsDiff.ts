// `sheets-framework configs-diff`: summarises how the generated config files on disk differ from their HEAD versions. Read-only; no credential.
import type { ValueConfigBase } from "../src/01_SpreadsheetSchema/configReaders/makeConfigs.ts";
import { configFiles } from "./configsDiff/configFiles.ts";
import { snapshotDiff } from "./configsDiff/snapshotDiff.ts";
import { PackageConfigLoader } from "./PackageConfigLoader.ts";
import type { SheetsConfig } from "./sheetsConfig.ts";

export type ConfigEntry = Record<string, unknown>;
export type IdField = "tableId" | "columnId";
export type IdentifiedEntry<FN extends IdField> = ConfigEntry &
  Record<FN, string>;
export type TabEntry = ConfigEntry & { sheetGid: number };
export type RuleGroups = Record<string, Record<string, ConfigEntry>>;

export interface RuleSnapshot {
  tables: RuleGroups;
  tabs: RuleGroups;
}

export interface ConfigSnapshot {
  tableConfigs: Record<string, IdentifiedEntry<"tableId">>;
  tabConfigs: Record<string, TabEntry>;
  retiredIdPrefixes: string[];
  columnConfigs: Record<string, Record<string, IdentifiedEntry<"columnId">>>;
  valueConfigs: Record<string, ValueConfigBase>;
  conditionalFormatConfigs: RuleSnapshot;
}

export interface SnapshotPair {
  before: ConfigSnapshot | undefined;
  after: ConfigSnapshot;
}

export function summarizeConfigsDiff({
  before = emptySnapshot(),
  after,
}: SnapshotPair): string[] {
  return snapshotDiff.summarize({ before, after });
}

function emptySnapshot(): ConfigSnapshot {
  return {
    tableConfigs: {},
    tabConfigs: {},
    retiredIdPrefixes: [],
    columnConfigs: {},
    valueConfigs: {},
    conditionalFormatConfigs: { tables: {}, tabs: {} },
  };
}

export function runConfigsDiff(sheetsConfig: SheetsConfig): void {
  if (!PackageConfigLoader.fromSheetsConfig(sheetsConfig).hasConfigs()) {
    throw new Error(
      `configs:diff: no generated config files in ${sheetsConfig.generatedDir}; run sync first.`,
    );
  }
  console.log(readConfigsDiff(sheetsConfig).join("\n"));
}

export function readConfigsDiff(sheetsConfig: SheetsConfig): string[] {
  const { dir } = sheetsConfig;
  const configLoader = PackageConfigLoader.fromSheetsConfig(sheetsConfig);
  return summarizeConfigsDiff({
    before: configFiles.readSnapshot((file) =>
      configFiles.readHeadFile(dir, configLoader.filePath(file)),
    ),
    after: configFiles.readSnapshot((file) =>
      configFiles.readFileIfPresent(configLoader.filePath(file)),
    ),
  });
}
