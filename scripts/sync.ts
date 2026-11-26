// `sheets-framework sync`: runs sync on the Node host, planning against the live Tables and previous configs, and sending only with --send.
import { spawnSync } from "node:child_process";
import { existsSync, mkdirSync, writeFileSync } from "node:fs";
import { basename } from "node:path";

import type {
  ForcedTrait,
  SyncRunResult,
} from "../src/01_SpreadsheetSchema/configReaders/syncRunResult.ts";
import type {
  ConfigRegeneration,
  SyncRun,
} from "../src/05_Operators/SyncCoordinator.ts";
import type { HeldList } from "../src/nodeHost/EnteredDataGate.ts";
import { readConfigsDiff } from "./configsDiff.ts";
import { GeneratedDir } from "./GeneratedDir.ts";
import { HeldFile, parseConfirmToken } from "./HeldFile.ts";
import {
  hasPackageConfigs,
  loadDropdownStyleReset,
  loadFrameworkConfigs,
  loadPackageConfigs,
  startNodeHost,
} from "./nodeHost.ts";
import type { SheetsConfig } from "./sheetsConfig.ts";

interface SyncRunnerProps {
  sheetsConfig: SheetsConfig;
  isSending: boolean;
  forcedTraits: ForcedTrait[];
  confirmToken?: string;
}

export class SyncRunner {
  readonly sheetsConfig: SheetsConfig;
  readonly isSending: boolean;
  readonly forcedTraits: ForcedTrait[];
  readonly heldFile: HeldFile;
  readonly generatedDir: GeneratedDir;
  constructor({
    sheetsConfig,
    isSending,
    forcedTraits,
    confirmToken,
  }: SyncRunnerProps) {
    this.sheetsConfig = sheetsConfig;
    this.isSending = isSending;
    this.forcedTraits = forcedTraits;
    this.heldFile = HeldFile.init({
      packageDir: sheetsConfig.dir,
      command: "sync",
      isSending,
      confirmToken,
    });
    this.generatedDir = new GeneratedDir(sheetsConfig.generatedDir);
  }
  static init(sheetsConfig: SheetsConfig, argv: string[]): SyncRunner {
    return new SyncRunner({
      sheetsConfig,
      isSending: argv.includes("--send"),
      forcedTraits: parseForcedTraits(argv),
      confirmToken: parseConfirmToken(argv),
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
    console.log(`\n${labelled("changes against HEAD:")}`);
    console.log(readConfigsDiff(this.sheetsConfig).join("\n"));
    console.log(`\n${report}`);

    console.log(
      "\nRunning this package's npm run tsc to check the regenerated files...",
    );
    if (!this._runTsc()) {
      reportTscFailure();
      process.exitCode = 1;
      return;
    }
    console.log(labelled("tsc passed."));
  }

  async _sync(): Promise<SyncRun> {
    const { generatedDir } = this;
    const target = basename(this.sheetsConfig.dir);
    // Loaded before the host starts, as the configs are, so a broken alterations.ts costs no setup.
    const intendedState = await generatedDir.loadIntendedState();
    const previousIdNameMap = generatedDir.loadPreviousIdNameMap();
    // The previous configs say which Tables have API access; a package with no generated files yet borrows the framework's.
    // No app endpoints: they may name columns the configs being regenerated don't have yet.
    const host = await startNodeHost({
      isDryRun: !this.isSending,
      sheetsConfig: this.sheetsConfig,
      configs: hasPackageConfigs(this.sheetsConfig)
        ? await loadPackageConfigs(this.sheetsConfig)
        : await loadFrameworkConfigs(),
      enteredDataGate: {
        target,
        confirmedList: this.heldFile.confirmedList,
        forcedTraits: this.forcedTraits,
      },
    });
    const { SyncCoordinator } =
      await import("../src/05_Operators/SyncCoordinator.ts");
    const run = SyncCoordinator.init().run({
      isSending: this.isSending,
      intendedState,
      previousIdNameMap,
      forcedTraits: this.forcedTraits,
      frame: {
        target,
        command: "npm run sync",
        generatedDir: generatedDir.dir,
        alterationsPath: generatedDir.path.alterations,
      },
      isHostWithholding: true,
      dropdownStyleReset: await loadDropdownStyleReset(this.sheetsConfig),
    });
    const { heldList } = host;
    if (heldList === undefined) return run;
    return { ...run, result: await this._withHeld(run.result, heldList) };
  }

  // Only a dry run prints the held list and its token, and writes every held item with its value to the held file.
  async _withHeld(
    result: SyncRunResult,
    heldList: HeldList,
  ): Promise<SyncRunResult> {
    if (result.mode !== "dry run") return result;
    const held = await this.heldFile.dryRunHeld(heldList, result.target);
    if (held === undefined) return result;
    return { ...result, held };
  }

  // Write nothing until every config and the ID-name map are confirmed good; a subset would go stale.
  _write(files: ConfigRegeneration): void {
    const { dir, path } = this.generatedDir;
    mkdirSync(dir, { recursive: true });
    const writes: [string, string][] = [
      [path.tableConfigs, files.tableConfigs],
      [path.tabConfigs, files.tabConfigs],
      [path.conditionalFormatConfigs, files.conditionalFormatConfigs],
      [path.retiredIdPrefixes, files.retiredIdPrefixes],
      [path.columnConfigs, files.columnConfigs],
      [path.valueConfigs, files.valueConfigs],
      [path.appConfigs, GeneratedDir.appConfigsText],
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
      writeFileSync(path.alterations, GeneratedDir.alterationsText);
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

// Each `--force <table>/<column>/<trait>`, for this run only: nothing carries it to the next.
export function parseForcedTraits(argv: string[]): ForcedTrait[] {
  return argv.flatMap((arg, index): ForcedTrait[] => {
    if (arg !== "--force") return [];
    const address = argv[index + 1] ?? "";
    const [table, column, trait, ...rest] = address.split("/");
    if (!table || !column || !trait || rest.length > 0) {
      throw new Error(
        `--force takes <table>/<column>/<trait>, got "${address}".`,
      );
    }
    return [{ table, column, trait }];
  });
}

function labelled(text: string): string {
  return `sync: ${text}`;
}

function reportTscFailure(): void {
  console.error(
    `\n${labelled("regeneration succeeded and every generated file was written, ")}` +
      "but this package's `npm run tsc` failed above. This usually means " +
      "hand-written references in this package still name a sheet or column " +
      "that no longer exists after this regeneration. Fix those references " +
      "and re-run `npm run tsc` — do not hand-edit the generated files.",
  );
}

// Non-zero whenever something the run meant to do didn't happen, so an agent can branch on the exit code alone.
function syncExitCode(result: SyncRunResult): number {
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
  await SyncRunner.init(sheetsConfig, argv).run();
}
