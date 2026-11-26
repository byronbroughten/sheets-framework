// `sheets-framework chore`: runs one chore against the package's spreadsheet. See docs/how-it-runs/chore-dry-run.md.
import { fileURLToPath } from "node:url";

import type { SheetsHttpTransport } from "../src/00_Source/GoogleSheets/GoogleSheetsAPI.ts";
import type { HeldFlushRefusal } from "../src/00_Source/RawSource/HeldFlushRefusal.ts";
import type { Configs } from "../src/01_SpreadsheetSchema/configReaders/configRegister.ts";
import type { DryRunHeld } from "../src/01_SpreadsheetSchema/configReaders/dryRunHeld.ts";
import type { Chore } from "../src/chores/Chore.ts";
import type { NodeHost } from "../src/nodeHost/NodeHost.ts";
import { ChoreIndex } from "./choreIndex.ts";
import { HeldFile, parseConfirmToken } from "./HeldFile.ts";
import {
  importConstNamedAfterFile,
  loadPackageConfigs,
  loadPackageEndpoints,
  startNodeHost,
} from "./nodeHost.ts";
import type { SheetsConfig } from "./sheetsConfig.ts";
import {
  snapshotFlags,
  SnapshotFolder,
  snapshotTarget,
} from "./SnapshotFolder.ts";

// Flags whose next argument is their value, not a chore's name.
const valueFlags: string[] = ["--confirm", ...Object.values(snapshotFlags)];

const genericHome = fileURLToPath(new URL("../src/chores/", import.meta.url));

// A test's stand-ins for Google, for the package's configs, which a program installs once, and for the clock.
export interface ChoreHostStandIn {
  transport: SheetsHttpTransport;
  configs: Configs;
  now?: () => Date;
}

interface ChoreRunnerProps {
  sheetsConfig: SheetsConfig;
  index: ChoreIndex;
  choreName: string | undefined;
  isSending: boolean;
  isJson: boolean;
  heldFile: HeldFile;
  snapshotFolder: SnapshotFolder;
  standIn: ChoreHostStandIn | undefined;
}

export class ChoreRunner {
  readonly sheetsConfig: SheetsConfig;
  readonly index: ChoreIndex;
  readonly choreName: string | undefined;
  readonly isSending: boolean;
  readonly isJson: boolean;
  readonly heldFile: HeldFile;
  readonly snapshotFolder: SnapshotFolder;
  readonly standIn: ChoreHostStandIn | undefined;
  constructor({
    sheetsConfig,
    index,
    choreName,
    isSending,
    isJson,
    heldFile,
    snapshotFolder,
    standIn,
  }: ChoreRunnerProps) {
    this.sheetsConfig = sheetsConfig;
    this.index = index;
    this.choreName = choreName;
    this.isSending = isSending;
    this.isJson = isJson;
    this.heldFile = heldFile;
    this.snapshotFolder = snapshotFolder;
    this.standIn = standIn;
  }
  static init(
    sheetsConfig: SheetsConfig,
    argv: string[],
    standIn?: ChoreHostStandIn,
  ): ChoreRunner {
    const confirmToken = parseConfirmToken(argv);
    const isSending = argv.includes("--send");
    return new ChoreRunner({
      sheetsConfig,
      index: ChoreIndex.init({
        genericHome,
        packageHomes: sheetsConfig.choreHomes,
      }),
      // A flag's value, the token after --confirm among them, is no chore name.
      choreName: argv.find(
        (arg, index) =>
          !arg.startsWith("--") && !valueFlags.includes(argv[index - 1] ?? ""),
      ),
      isSending,
      isJson: argv.includes("--json"),
      heldFile: HeldFile.init({
        packageDir: sheetsConfig.dir,
        command: "chore",
        isSending,
        confirmToken,
      }),
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
  async run(): Promise<void> {
    if (!this.choreName) {
      console.log(
        "Usage: npm run <app|dev>:chore <name> [-- --send [--confirm <token>]] [-- --json]",
      );
      console.log(
        "       a --confirm send also takes [--snapshots <folder>] [--restore-script <name>]",
      );
      console.log("  no flag    preview what it would write, writing nothing");
      console.log("  --send     apply it to the live spreadsheet");
      console.log(
        "  --confirm  with --send, send what the dry run held, by its token",
      );
      console.log("  --json     preview as raw request JSON");
      console.log(
        "  --snapshots  with --confirm, the folder the snapshot file goes in (default .probe/snapshots)",
      );
      console.log(
        "  --restore-script  with --confirm, the npm script the undo line names (default restore)\n",
      );
      console.log(this.index.listing(this.sheetsConfig.dir));
      return;
    }
    // Resolved before the host starts, so a typo costs no setup.
    const modulePath = this._choreModulePath(this.choreName);
    const host = await this._startHost();
    const chore = await importConstNamedAfterFile<Chore>(
      modulePath,
      "A chore file",
    );
    console.log(`chore: ${this.choreName} — ${chore.description}\n`);
    const { SpreadsheetNamed } =
      await import("../src/04_SpreadsheetNamed/SpreadsheetNamed.ts");
    const { HeldFlushRefusal } =
      await import("../src/00_Source/RawSource/HeldFlushRefusal.ts");
    let result: string | void;
    try {
      result = chore.action(SpreadsheetNamed.init(), {
        spreadsheetId: host.spreadsheetId,
      });
    } catch (error) {
      // Flushes that landed before the stop still get their snapshot.
      const isRefusal = error instanceof HeldFlushRefusal;
      if (isRefusal) await this._reportRefusal(error);
      await this._keepSnapshot(host);
      if (!isRefusal) throw error;
      return;
    }
    if (result) console.log(`result: ${result}`);
    const held = await this._dryRunHeld(host);
    this._report(host, held);
    if (held !== undefined) await reportHeld(held);
    await this._keepSnapshot(host);
  }
  // A confirmed send prunes the old snapshots and writes its own, then prints the undo line; a send nothing held has no snapshot.
  async _keepSnapshot(host: NodeHost): Promise<void> {
    if (!this.isSending || this.heldFile.confirmToken === undefined) return;
    const { snapshotFolder } = this;
    snapshotFolder.prune();
    const snapshot = host.sendSnapshot();
    if (snapshot === undefined) return;
    const path = snapshotFolder.write({
      target: this.target,
      command: `chore ${this.choreName}`,
      plan: this._planText(host),
      snapshot,
    });
    const lines = await snapshotFolder.undoLines(
      path,
      snapshot.versionHistoryItems,
    );
    console.log(`\n${lines.join("\n")}`);
  }
  async _startHost(): Promise<NodeHost> {
    const { standIn } = this;
    return startNodeHost({
      isDryRun: !this.isSending,
      sheetsConfig: this.sheetsConfig,
      configs:
        standIn?.configs ?? (await loadPackageConfigs(this.sheetsConfig)),
      endpoints: await loadPackageEndpoints(this.sheetsConfig),
      enteredDataGate: {
        target: this.target,
        confirmedList: this.heldFile.confirmedList,
      },
      transport: standIn?.transport,
    });
  }
  _choreModulePath(choreName: string): string {
    const path = this.index.pathOf(choreName);
    if (!path) {
      throw new Error(
        `No chore named "${choreName}".\n\n${this.index.listing(this.sheetsConfig.dir)}`,
      );
    }
    return path;
  }
  // Only a dry run prints the held list and its token, and writes every held item with its value to the held file.
  async _dryRunHeld({ heldList }: NodeHost): Promise<DryRunHeld | undefined> {
    if (this.isSending || heldList === undefined) return undefined;
    return this.heldFile.dryRunHeld(heldList, this.target);
  }
  _report({ summary }: NodeHost, held: DryRunHeld | undefined): void {
    if (summary.isEmpty) {
      console.log("\nThis chore wrote nothing.");
      return;
    }
    const heading = this.isSending
      ? "SENT — the spreadsheet was written to:"
      : "DRY RUN — nothing was written. It would send:";
    console.log(`\n${heading}\n`);
    console.log(this._planText({ summary }));
    console.log(`\n${summary.count} request(s).${this._nextStep(held)}`);
  }
  // The requests as the run prints them, which the send's snapshot file keeps as its plan.
  _planText({ summary }: Pick<NodeHost, "summary">): string {
    return this.isJson ? summary.json : summary.lines.join("\n");
  }
  // A dry run that held items points at the held section's rerun instead, since a plain --send would be refused.
  _nextStep(held: DryRunHeld | undefined): string {
    if (this.isSending || held !== undefined) return "";
    return " Re-run with `-- --send` to apply.";
  }
  async _reportRefusal(refusal: HeldFlushRefusal): Promise<void> {
    const { heldStoppedOutcome } =
      await import("../src/01_SpreadsheetSchema/configReaders/flushStoppedOutcome.ts");
    const { GateReport } = await import("../src/nodeHost/GateReport.ts");
    console.log(
      `\n${GateReport.flushStoppedLine(heldStoppedOutcome(refusal))}`,
    );
    console.log(GateReport.rerunDryRunLine(`npm run chore ${this.choreName}`));
    process.exitCode = 1;
  }
}

async function reportHeld(held: DryRunHeld): Promise<void> {
  const { GateReport } = await import("../src/nodeHost/GateReport.ts");
  console.log(GateReport.heldSection(held).join("\n"));
}

export async function runChore(
  sheetsConfig: SheetsConfig,
  argv: string[],
): Promise<void> {
  await ChoreRunner.init(sheetsConfig, argv).run();
}
