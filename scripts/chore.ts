// `sheets-framework chore`: runs one chore against the package's spreadsheet. See docs/how-it-runs/chore-dry-run.md.
import { basename } from "node:path";
import { fileURLToPath } from "node:url";

import type { HeldFlushRefusal } from "../src/00_Source/RawSource/HeldFlushRefusal.ts";
import type {
  SyncStoppedBeforeSendingOutcome,
  SyncStoppedPartwayOutcome,
} from "../src/01_SpreadsheetSchema/configReaders/syncRunResult.ts";
import type { Chore } from "../src/chores/Chore.ts";
import type { ConfirmedList } from "../src/nodeHost/EnteredDataGate.ts";
import type { NodeHost } from "../src/nodeHost/NodeHost.ts";
import { ChoreIndex } from "./choreIndex.ts";
import { heldFile, parseConfirmToken } from "./heldFile.ts";
import {
  importConstNamedAfterFile,
  loadPackageConfigs,
  loadPackageEndpoints,
  startNodeHost,
} from "./nodeHost.ts";
import type { SheetsConfig } from "./sheetsConfig.ts";

const genericHome = fileURLToPath(new URL("../src/chores/", import.meta.url));

interface ChoreRunnerProps {
  sheetsConfig: SheetsConfig;
  index: ChoreIndex;
  choreName: string | undefined;
  isSend: boolean;
  isJson: boolean;
  confirmToken: string | undefined;
}

export class ChoreRunner {
  readonly sheetsConfig: SheetsConfig;
  readonly index: ChoreIndex;
  readonly choreName: string | undefined;
  readonly isSend: boolean;
  readonly isJson: boolean;
  readonly confirmToken: string | undefined;
  readonly heldFilePath: string;
  constructor({
    sheetsConfig,
    index,
    choreName,
    isSend,
    isJson,
    confirmToken,
  }: ChoreRunnerProps) {
    this.sheetsConfig = sheetsConfig;
    this.index = index;
    this.choreName = choreName;
    this.isSend = isSend;
    this.isJson = isJson;
    this.confirmToken = confirmToken;
    this.heldFilePath = heldFile.path(sheetsConfig.dir, "chore");
  }
  static init(sheetsConfig: SheetsConfig, argv: string[]): ChoreRunner {
    const confirmToken = parseConfirmToken(argv);
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
      isSend: argv.includes("--send"),
      isJson: argv.includes("--json"),
      confirmToken,
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
    this._report(host);
    await this._reportHeld(host);
  }
  async _startHost(): Promise<NodeHost> {
    return startNodeHost({
      isDryRun: !this.isSend,
      sheetsConfig: this.sheetsConfig,
      configs: await loadPackageConfigs(this.sheetsConfig),
      endpoints: await loadPackageEndpoints(this.sheetsConfig),
      enteredDataGate: {
        target: this.target,
        confirmedList: this._confirmedList(),
      },
    });
  }
  _confirmedList(): ConfirmedList | undefined {
    return heldFile.confirmedList({
      path: this.heldFilePath,
      isSending: this.isSend,
      confirmToken: this.confirmToken,
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
  _report({ summary }: NodeHost): void {
    if (summary.isEmpty) {
      console.log("\nThis chore wrote nothing.");
      return;
    }
    const heading = this.isSend
      ? "SENT — the spreadsheet was written to:"
      : "DRY RUN — nothing was written. It would send:";
    console.log(`\n${heading}\n`);
    console.log(this.isJson ? summary.json : summary.lines.join("\n"));
    const nextStep = this.isSend ? "" : " Re-run with `-- --send` to apply.";
    console.log(`\n${summary.count} request(s).${nextStep}`);
  }
  // Only a dry run prints the held list and its token, and writes every held item with its value to the held file.
  async _reportHeld({ heldList }: NodeHost): Promise<void> {
    if (this.isSend || heldList === undefined) return;
    const held = await heldFile.dryRunHeld({
      path: this.heldFilePath,
      target: this.target,
      heldList,
    });
    if (held === undefined) return;
    const { heldSection } = await import("../src/nodeHost/SyncRunReport.ts");
    console.log(heldSection(held).join("\n"));
  }
  async _reportRefusal(refusal: HeldFlushRefusal): Promise<void> {
    const { heldRefusalLine, rerunDryRunLine } =
      await import("../src/nodeHost/SyncRunReport.ts");
    console.log(`\n${heldRefusalLine(refusalOutcome(refusal))}`);
    console.log(rerunDryRunLine(`npm run chore ${this.choreName}`));
    process.exitCode = 1;
  }
}

// The gate's refusal in the shape sync's report reads; one refused at the first flush landed nothing.
function refusalOutcome({
  kind,
  landedFlushes,
  flushCount,
}: HeldFlushRefusal):
  SyncStoppedBeforeSendingOutcome | SyncStoppedPartwayOutcome {
  if (landedFlushes === 0) {
    return { kind: "stoppedBeforeSending", heldBy: kind };
  }
  return { kind: "stoppedPartway", landedFlushes, flushCount, heldBy: kind };
}

export async function runChore(
  sheetsConfig: SheetsConfig,
  argv: string[],
): Promise<void> {
  await ChoreRunner.init(sheetsConfig, argv).run();
}
