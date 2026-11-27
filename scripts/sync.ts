// `sheets-framework sync`: runs sync on the Node host, planning against the live Tables and previous configs, and sending only with --send.
import { spawnSync } from "node:child_process";
import { existsSync, mkdirSync, writeFileSync } from "node:fs";

import type { Configs } from "../src/01_SpreadsheetSchema/configReaders/configRegister.ts";
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
import { type HostStandIn, startNodeHost } from "./nodeHost.ts";
import { PackageConfigLoader } from "./PackageConfigLoader.ts";
import {
  loadDropdownStyleReset,
  loadFrameworkConfigs,
} from "./packageLoader.ts";
import type { SheetsConfig } from "./sheetsConfig.ts";
import {
  printSnapshotNotice,
  SnapshotFolder,
  snapshotTarget,
} from "./SnapshotFolder.ts";
import { labelled } from "./syncLabel.ts";
import {
  type SnapshotNotice,
  SyncSnapshotKeeper,
} from "./SyncSnapshotKeeper.ts";

// A test's stand-ins: the host's, and the package's tsc, which a temporary package can't run.
export interface SyncStandIn extends HostStandIn {
  runTsc: () => boolean;
}

interface SyncRunnerProps {
  sheetsConfig: SheetsConfig;
  isSending: boolean;
  forcedTraits: ForcedTrait[];
  confirmToken?: string;
  snapshotFolder: SnapshotFolder;
  standIn: SyncStandIn | undefined;
}

// A run with what its held send says of its snapshot, when it says anything.
interface KeptSyncRun extends SyncRun {
  snapshotNotice?: SnapshotNotice;
}

export class SyncRunner {
  readonly sheetsConfig: SheetsConfig;
  readonly isSending: boolean;
  readonly forcedTraits: ForcedTrait[];
  readonly heldFile: HeldFile;
  readonly generatedDir: GeneratedDir;
  readonly snapshotFolder: SnapshotFolder;
  readonly standIn: SyncStandIn | undefined;
  constructor({
    sheetsConfig,
    isSending,
    forcedTraits,
    confirmToken,
    snapshotFolder,
    standIn,
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
    this.snapshotFolder = snapshotFolder;
    this.standIn = standIn;
  }
  static init(
    sheetsConfig: SheetsConfig,
    argv: string[],
    standIn?: SyncStandIn,
  ): SyncRunner {
    return new SyncRunner({
      sheetsConfig,
      isSending: argv.includes("--send"),
      forcedTraits: parseForcedTraits(argv),
      confirmToken: parseConfirmToken(argv),
      snapshotFolder: SnapshotFolder.fromArgv({
        packageDir: sheetsConfig.dir,
        argv,
        now: standIn?.now ?? (() => new Date()),
      }),
      standIn,
    });
  }
  get target(): string {
    return snapshotTarget(this.sheetsConfig);
  }
  // A confirmed or forced send.
  get isHeldSend(): boolean {
    const isConfirmed = this.heldFile.confirmToken !== undefined;
    const isForced = this.forcedTraits.length > 0;
    return this.isSending && (isConfirmed || isForced);
  }
  async run(): Promise<void> {
    const synced = await this._sync();
    const { result, files, snapshotNotice } = synced;
    const report = await reportText(result);
    process.exitCode = syncExitCode(synced);
    // A refused run returns no files: the run writes every file or none.
    if (files === undefined) {
      console.log(report);
      printSnapshotNotice(snapshotNotice?.text);
      return;
    }
    try {
      this._write(files);
      console.log(`\n${labelled("changes against HEAD:")}`);
      console.log(readConfigsDiff(this.sheetsConfig).join("\n"));
      console.log(`\n${report}`);
    } finally {
      // The send is already out, so a failing write or diff still leaves its undo line printed.
      printSnapshotNotice(snapshotNotice?.text);
    }

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

  async _sync(): Promise<KeptSyncRun> {
    const { generatedDir, standIn, target } = this;
    // Loaded before the host starts, as the configs are, so a broken alterations.ts costs no setup.
    const intendedState = await generatedDir.loadIntendedState();
    const previousIdNameMap = generatedDir.loadPreviousIdNameMap();
    // The previous configs say which Tables have API access.
    // No app endpoints: they may name columns the configs being regenerated don't have yet.
    const host = await startNodeHost({
      isDryRun: !this.isSending,
      sheetsConfig: this.sheetsConfig,
      configs: standIn?.configs ?? (await this._loadConfigs()),
      enteredDataGate: {
        target,
        confirmedList: this.heldFile.confirmedList,
        forcedTraits: this.forcedTraits,
      },
      transport: standIn?.transport,
    });
    const { SyncCoordinator } =
      await import("../src/05_Operators/SyncCoordinator.ts");
    const dropdownStyleReset = await loadDropdownStyleReset(this.sheetsConfig);
    const keeper = SyncSnapshotKeeper.init({
      host,
      snapshotFolder: this.snapshotFolder,
      target,
      isHeldSend: this.isHeldSend,
    });
    let run: SyncRun;
    try {
      run = SyncCoordinator.init().run({
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
        dropdownStyleReset,
      });
    } catch (error) {
      await keeper.keepThrown();
      throw error;
    }
    // Kept before any generated file is written, so a failing write leaves the snapshot behind.
    const snapshotNotice = await keeper.keepReturned({
      result: run.result,
      plan: await reportText(run.result),
    });
    const { heldList } = host;
    if (heldList === undefined) return { ...run, snapshotNotice };
    const result = await this._withHeld(run.result, heldList);
    return { ...run, result, snapshotNotice };
  }
  // A package with no generated files yet borrows the framework's.
  private async _loadConfigs(): Promise<Configs> {
    const packageConfigs = PackageConfigLoader.fromSheetsConfig(
      this.sheetsConfig,
    );
    if (packageConfigs.hasConfigs()) {
      return packageConfigs.loadConfigs();
    }
    return loadFrameworkConfigs();
  }

  // Only a dry run prints the held list and its token, and writes every held item with its value to the held file.
  async _withHeld(
    result: SyncRunResult,
    heldList: HeldList,
  ): Promise<SyncRunResult> {
    if (result.mode !== "dry run") return result;
    // The argv's traits, the list the gate's send snapshot names.
    const held = await this.heldFile.dryRunHeld({
      heldList,
      target: result.target,
      forcedTraits: this.forcedTraits,
    });
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
    if (this.standIn !== undefined) return this.standIn.runTsc();
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

// The report as the run prints it, which the send's snapshot file keeps as its plan.
async function reportText(result: SyncRunResult): Promise<string> {
  const { SyncRunReport } = await import("../src/nodeHost/SyncRunReport.ts");
  return SyncRunReport.init(result).text;
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
function syncExitCode({ result, snapshotNotice }: KeptSyncRun): number {
  const { outcome } = result;
  if (snapshotNotice?.isLost) return 1;
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
