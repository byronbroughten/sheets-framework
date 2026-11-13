// `sheets-framework chore`: runs one chore against the package's spreadsheet. See docs/how-it-runs/chore-dry-run.md.
import { basename } from "node:path";
import { fileURLToPath } from "node:url";

import type { SheetsHttpTransport } from "../src/00_Source/GoogleSheets/GoogleSheetsAPI.ts";
import type { HeldFlushRefusal } from "../src/00_Source/RawSource/HeldFlushRefusal.ts";
import type { Configs } from "../src/01_SpreadsheetSchema/configReaders/configRegister.ts";
import type { SyncHeld } from "../src/01_SpreadsheetSchema/configReaders/syncRunResult.ts";
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

const genericHome = fileURLToPath(new URL("../src/chores/", import.meta.url));

// A test's stand-ins for Google and for the package's configs, which a program installs once.
export interface ChoreHostStandIn {
  transport: SheetsHttpTransport;
  configs: Configs;
}

interface ChoreRunnerProps {
  sheetsConfig: SheetsConfig;
  index: ChoreIndex;
  choreName: string | undefined;
  isSending: boolean;
  isJson: boolean;
  heldFile: HeldFile;
  standIn: ChoreHostStandIn | undefined;
}

export class ChoreRunner {
  readonly sheetsConfig: SheetsConfig;
  readonly index: ChoreIndex;
  readonly choreName: string | undefined;
  readonly isSending: boolean;
  readonly isJson: boolean;
  readonly heldFile: HeldFile;
  readonly standIn: ChoreHostStandIn | undefined;
  constructor({
    sheetsConfig,
    index,
    choreName,
    isSending,
    isJson,
    heldFile,
    standIn,
  }: ChoreRunnerProps) {
    this.sheetsConfig = sheetsConfig;
    this.index = index;
    this.choreName = choreName;
    this.isSending = isSending;
    this.isJson = isJson;
    this.heldFile = heldFile;
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
      // The token after --confirm is no chore name.
      choreName: argv.find(
        (arg) => !arg.startsWith("--") && arg !== confirmToken,
      ),
      isSending,
      isJson: argv.includes("--json"),
      heldFile: HeldFile.init({
        packageDir: sheetsConfig.dir,
        command: "chore",
        isSending,
        confirmToken,
      }),
      standIn,
    });
  }
  get target(): string {
    return basename(this.sheetsConfig.dir);
  }
  async run(): Promise<void> {
    if (!this.choreName) {
      console.log(
        "Usage: npm run <app|dev>:chore <name> [-- --send [--confirm <token>]] [-- --json]",
      );
      console.log("  no flag    preview what it would write, writing nothing");
      console.log("  --send     apply it to the live spreadsheet");
      console.log(
        "  --confirm  with --send, send what the dry run held, by its token",
      );
      console.log("  --json     preview as raw request JSON\n");
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
      if (!(error instanceof HeldFlushRefusal)) throw error;
      await this._reportRefusal(error);
      return;
    }
    if (result) console.log(`result: ${result}`);
    const held = await this._dryRunHeld(host);
    this._report(host, held);
    if (held !== undefined) await reportHeld(held);
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
  async _dryRunHeld({ heldList }: NodeHost): Promise<SyncHeld | undefined> {
    if (this.isSending || heldList === undefined) return undefined;
    return this.heldFile.dryRunHeld(heldList, this.target);
  }
  _report({ summary }: NodeHost, held: SyncHeld | undefined): void {
    if (summary.isEmpty) {
      console.log("\nThis chore wrote nothing.");
      return;
    }
    const heading = this.isSending
      ? "SENT — the spreadsheet was written to:"
      : "DRY RUN — nothing was written. It would send:";
    console.log(`\n${heading}\n`);
    console.log(this.isJson ? summary.json : summary.lines.join("\n"));
    console.log(`\n${summary.count} request(s).${this._nextStep(held)}`);
  }
  // A dry run that held items points at the held section's rerun instead, since a plain --send would be refused.
  _nextStep(held: SyncHeld | undefined): string {
    if (this.isSending || held !== undefined) return "";
    return " Re-run with `-- --send` to apply.";
  }
  async _reportRefusal(refusal: HeldFlushRefusal): Promise<void> {
    const { heldStoppedOutcome } =
      await import("../src/01_SpreadsheetSchema/configReaders/heldStoppedOutcome.ts");
    const { heldRefusalLine, rerunDryRunLine } =
      await import("../src/nodeHost/SyncRunReport.ts");
    console.log(`\n${heldRefusalLine(heldStoppedOutcome(refusal))}`);
    console.log(rerunDryRunLine(`npm run chore ${this.choreName}`));
    process.exitCode = 1;
  }
}

async function reportHeld(held: SyncHeld): Promise<void> {
  const { heldSection } = await import("../src/nodeHost/SyncRunReport.ts");
  console.log(heldSection(held).join("\n"));
}

export async function runChore(
  sheetsConfig: SheetsConfig,
  argv: string[],
): Promise<void> {
  await ChoreRunner.init(sheetsConfig, argv).run();
}
