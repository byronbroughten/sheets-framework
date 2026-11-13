// `sheets-framework sync`: runs sync on the Node host, planning against the live Tables and previous configs, and sending only with --send.
import { spawnSync } from "node:child_process";
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { basename, dirname, join } from "node:path";

import type { IdNameMapGeneric } from "../src/01_SpreadsheetSchema/configReaders/idNameMapTypes.ts";
import type { PendingAlterations } from "../src/01_SpreadsheetSchema/configReaders/intendedState.ts";
import type { SyncRunResult } from "../src/01_SpreadsheetSchema/configReaders/syncRunResult.ts";
import { alterationsFileSource } from "../src/05_Operators/configFileSource.ts";
import { parseIdNameMapFileSource } from "../src/05_Operators/parseIdNameMapFileSource.ts";
import type {
  ConfigRegeneration,
  SyncIntendedState,
  SyncRun,
} from "../src/05_Operators/SyncCoordinator.ts";
import { readConfigsDiff } from "./configsDiff.ts";
import {
  type ConfigFile,
  configFilePath,
  hasPackageConfigs,
  importExportIfPresent,
  loadFrameworkConfigs,
  loadPackageConfigs,
  startNodeHost,
} from "./nodeHost.ts";
import type { SheetsConfig } from "./sheetsConfig.ts";

interface ConfigFilesGeneratorProps {
  sheetsConfig: SheetsConfig;
  isSending: boolean;
}

class ConfigFilesGenerator {
  readonly sheetsConfig: SheetsConfig;
  readonly isSending: boolean;
  readonly path: Record<
    ConfigFile | "appConfigs" | "idNameMap" | "alterations",
    string
  >;
  constructor({ sheetsConfig, isSending }: ConfigFilesGeneratorProps) {
    this.sheetsConfig = sheetsConfig;
    this.isSending = isSending;
    const { generatedDir } = sheetsConfig;
    this.path = {
      tableConfigs: configFilePath(generatedDir, "tableConfigs"),
      tabConfigs: configFilePath(generatedDir, "tabConfigs"),
      conditionalFormatConfigs: configFilePath(
        generatedDir,
        "conditionalFormatConfigs",
      ),
      retiredIdPrefixes: configFilePath(generatedDir, "retiredIdPrefixes"),
      columnConfigs: configFilePath(generatedDir, "columnConfigs"),
      valueConfigs: configFilePath(generatedDir, "valueConfigs"),
      appConfigs: join(generatedDir, "appConfigs.ts"),
      idNameMap: idNameMapPath(generatedDir),
      alterations: alterationsPath(generatedDir),
    };
  }
  static init(
    sheetsConfig: SheetsConfig,
    argv: string[],
  ): ConfigFilesGenerator {
    return new ConfigFilesGenerator({
      sheetsConfig,
      isSending: argv.includes("--send"),
    });
  }
  async run(): Promise<void> {
    const { result, files } = await this._sync();
    const { SyncRunReport } = await import("../src/nodeHost/SyncRunReport.ts");
    const report = SyncRunReport.init(result).text;
    process.exitCode = syncExitCode(result);
    // A refused run returns no files: the run writes every file or none.
    if (files === undefined) {
      console.log(report);
      return;
    }
    this._write(files);
    console.log("\nsync: changes against HEAD:");
    console.log(readConfigsDiff(this.sheetsConfig).join("\n"));
    console.log(`\n${report}`);

    console.log(
      "\nRunning this package's npm run tsc to check the regenerated files...",
    );
    if (!this._runTsc()) {
      reportTscFailure();
      process.exit(1);
    }
    console.log("sync: tsc passed.");
  }

  async _sync(): Promise<SyncRun> {
    const { generatedDir } = this.sheetsConfig;
    // Loaded before the host starts, as the configs are, so a broken alterations.ts costs no setup.
    const intendedState = await loadIntendedState(generatedDir);
    const previousIdNameMap = loadPreviousIdNameMap(generatedDir);
    // The previous configs say which Tables have API access; a package with no generated files yet borrows the framework's.
    // No app endpoints: they may name columns the configs being regenerated don't have yet.
    await startNodeHost({
      isDryRun: !this.isSending,
      sheetsConfig: this.sheetsConfig,
      configs: hasPackageConfigs(this.sheetsConfig)
        ? await loadPackageConfigs(this.sheetsConfig)
        : await loadFrameworkConfigs(),
    });
    const { SyncCoordinator } =
      await import("../src/05_Operators/SyncCoordinator.ts");
    return SyncCoordinator.init().run({
      isSending: this.isSending,
      intendedState,
      previousIdNameMap,
      frame: {
        target: basename(this.sheetsConfig.dir),
        command: "npm run sync",
        generatedDir,
        alterationsPath: this.path.alterations,
      },
    });
  }

  // Write nothing until every config and the ID-name map are confirmed good; a subset would go stale.
  _write(files: ConfigRegeneration): void {
    const { path } = this;
    mkdirSync(this.sheetsConfig.generatedDir, { recursive: true });
    const writes: [string, string][] = [
      [path.tableConfigs, files.tableConfigs],
      [path.tabConfigs, files.tabConfigs],
      [path.conditionalFormatConfigs, files.conditionalFormatConfigs],
      [path.retiredIdPrefixes, files.retiredIdPrefixes],
      [path.columnConfigs, files.columnConfigs],
      [path.valueConfigs, files.valueConfigs],
      [path.appConfigs, appConfigsText()],
      [path.idNameMap, files.idNameMap],
    ];
    // alterations.ts is rewritten only when the run returns it, so a run that landed nothing leaves its comments alone.
    if (files.alterations !== undefined) {
      writes.push([path.alterations, files.alterations]);
    }
    for (const [filePath, source] of writes) {
      writeFileSync(filePath, source);
      console.log(`Wrote ${filePath}`);
    }
    if (!existsSync(path.alterations)) {
      writeFileSync(path.alterations, alterationsText());
      console.log(`Created ${path.alterations}`);
    }
  }

  _runTsc(): boolean {
    const { status } = spawnSync("npm", ["run", "tsc"], {
      cwd: this.sheetsConfig.dir,
      stdio: "inherit",
    });
    return status === 0;
  }
}

function reportTscFailure(): void {
  console.error(
    "\nsync: regeneration succeeded and every generated file was written," +
      "but this package's `npm run tsc` failed above. This usually means " +
      "hand-written references in this package still name a sheet or column " +
      "that no longer exists after this regeneration. Fix those references " +
      "and re-run `npm run tsc` — do not hand-edit the generated files.",
  );
}

export function idNameMapPath(generatedDir: string): string {
  return join(generatedDir, "idNameMap.ts");
}

// A fixed path from the generated directory, so finding it needs no setting of its own.
export function alterationsPath(generatedDir: string): string {
  return join(dirname(generatedDir), "alterations.ts");
}

export function alterationsText(): string {
  return alterationsFileSource({});
}

// Intended state is the pending alterations alone until claims arrive; a package with no alterations.ts has none pending.
export async function loadIntendedState(
  generatedDir: string,
): Promise<SyncIntendedState> {
  const path = alterationsPath(generatedDir);
  const pendingAlterations = await importExportIfPresent<PendingAlterations>(
    path,
    "pendingAlterations",
  );
  return { alterations: pendingAlterations ?? {} };
}

// A package that has never synced has no map, so every Table is new to it.
export function loadPreviousIdNameMap(generatedDir: string): IdNameMapGeneric {
  const path = idNameMapPath(generatedDir);
  if (!existsSync(path)) return { tables: {}, tabs: {} };
  return parseIdNameMapFileSource(readFileSync(path, "utf8"));
}

export function appConfigsText(): string {
  return `// Without it, a program that reaches this file only by dynamic import rejects the augmentation (TS2664).
import type {} from "@byronbroughten/sheets-framework";

import { columnConfigs } from "./columnConfigs";
import { conditionalFormatConfigs } from "./conditionalFormatConfigs";
import type { IdNameMap } from "./idNameMap";
import { retiredIdPrefixes } from "./retiredIdPrefixes";
import { tabConfigs } from "./tabConfigs";
import { tableConfigs } from "./tableConfigs";
import { valueConfigs } from "./valueConfigs";

export const appConfigs = {
  tableConfigs,
  tabConfigs,
  conditionalFormatConfigs,
  retiredIdPrefixes,
  columnConfigs,
  valueConfigs,
};

declare module "@byronbroughten/sheets-framework" {
  interface Register {
    configs: typeof appConfigs;
    idNameMap: IdNameMap;
  }
}
`;
}

// Non-zero whenever something the run meant to do didn't happen, so an agent can branch on the exit code alone.
export function syncExitCode(result: SyncRunResult): number {
  const { outcome } = result;
  if (outcome.kind === "ready" || outcome.kind === "sent") {
    if (result.mode === "send" && result.didntLand.length > 0) return 1;
    return 0;
  }
  return 1;
}

export async function runSync(
  sheetsConfig: SheetsConfig,
  argv: string[],
): Promise<void> {
  await ConfigFilesGenerator.init(sheetsConfig, argv).run();
}
