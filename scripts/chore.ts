// `sheets-framework chore`: runs one chore against the package's spreadsheet. See docs/how-it-runs.md, "The chore and its dry run".
import { fileURLToPath, pathToFileURL } from "node:url";

import type { Chore } from "../src/chores/Chore.ts";
import type { NodeHost } from "../src/nodeHost/NodeHost.ts";
import { ChoreIndex } from "./choreIndex.ts";
import { loadPackageConfigs, startNodeHost } from "./nodeHost.ts";
import type { SheetsConfig } from "./sheetsConfig.ts";

const genericHome = fileURLToPath(new URL("../src/chores/", import.meta.url));

interface ChoreRunnerProps {
  sheetsConfig: SheetsConfig;
  index: ChoreIndex;
  choreName: string | undefined;
  isSend: boolean;
  isJson: boolean;
}

class ChoreRunner {
  readonly sheetsConfig: SheetsConfig;
  readonly index: ChoreIndex;
  readonly choreName: string | undefined;
  readonly isSend: boolean;
  readonly isJson: boolean;
  constructor({
    sheetsConfig,
    index,
    choreName,
    isSend,
    isJson,
  }: ChoreRunnerProps) {
    this.sheetsConfig = sheetsConfig;
    this.index = index;
    this.choreName = choreName;
    this.isSend = isSend;
    this.isJson = isJson;
  }
  static init(sheetsConfig: SheetsConfig, argv: string[]): ChoreRunner {
    return new ChoreRunner({
      sheetsConfig,
      index: ChoreIndex.init({
        genericHome,
        packageHomes: sheetsConfig.choreHomes,
      }),
      choreName: argv.find((arg) => !arg.startsWith("--")),
      isSend: argv.includes("--send"),
      isJson: argv.includes("--json"),
    });
  }
  async run(): Promise<void> {
    if (!this.choreName) {
      console.log(
        "Usage: npm run <app|dev>:chore <name> [-- --send] [-- --json]",
      );
      console.log("  no flag   preview what it would write, writing nothing");
      console.log("  --send    apply it to the live spreadsheet");
      console.log("  --json    preview as raw request JSON\n");
      console.log(this.index.listing(this.sheetsConfig.dir));
      return;
    }
    // Resolved before the host starts, so a typo costs no setup.
    const modulePath = this._choreModulePath(this.choreName);
    const host = await startNodeHost({
      isDryRun: !this.isSend,
      sheetsConfig: this.sheetsConfig,
      configs: await loadPackageConfigs(this.sheetsConfig),
    });
    const chore = await loadChore(modulePath, this.choreName);
    console.log(`chore: ${this.choreName} — ${chore.description}\n`);
    const { SpreadsheetNamed } =
      await import("../src/04_SpreadsheetNamed/SpreadsheetNamed.ts");
    const result = chore.action(SpreadsheetNamed.init(), {
      spreadsheetId: host.spreadsheetId,
    });
    if (result) console.log(`result: ${result}`);
    this._report(host);
  }
  _choreModulePath(choreName: string): string {
    const path = this.index.pathOf(choreName);
    if (!path) {
      throw new Error(
        `No chore named "${choreName}".\n\n${this.index.listing(this.sheetsConfig.dir)}`,
      );
    }
    return pathToFileURL(path).href;
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
}

async function loadChore(
  modulePath: string,
  choreName: string,
): Promise<Chore> {
  const chore = (await import(modulePath))[choreName];
  if (!chore) {
    throw new Error(
      `${modulePath} exports no "${choreName}". A chore file exports one const named after the file.`,
    );
  }
  return chore;
}

export async function runChore(
  sheetsConfig: SheetsConfig,
  argv: string[],
): Promise<void> {
  await ChoreRunner.init(sheetsConfig, argv).run();
}
