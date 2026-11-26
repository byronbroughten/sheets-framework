import { mkdirSync, mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

import { vi } from "vitest";

import { appConfigs } from "../dev/generated/appConfigs.ts";
import type { SheetsHttpTransport } from "../src/00_Source/GoogleSheets/GoogleSheetsAPI.ts";
import { ChoreRunner } from "./chore.ts";
import { RestoreRunner } from "./restore.ts";
import type { SheetsConfig } from "./sheetsConfig.ts";

interface ChoreSource {
  name: string;
  source: string;
}

// A temporary package whose one chore home holds the given chore.
export function choreSheetsConfig({ name, source }: ChoreSource): SheetsConfig {
  const packageDir = mkdtempSync(join(tmpdir(), "chore-"));
  const choreHome = join(packageDir, "chores");
  mkdirSync(choreHome);
  writeFileSync(join(choreHome, `${name}.ts`), source);
  return {
    path: join(packageDir, "sheets.config.json"),
    dir: packageDir,
    spreadsheetId: "spreadsheet-under-test",
    generatedDir: join(packageDir, "generated"),
    choreHomes: [choreHome],
    endpointModule: undefined,
  };
}

export interface ChoreRun {
  output: string;
  exitCode: number | string | null | undefined;
}

interface ChoreRunProps {
  sheetsConfig: SheetsConfig;
  transport: SheetsHttpTransport;
  argv: string[];
  now?: () => Date;
}

// Runs a chore past its live host: only the Sheets transport, and the clock when a test names one, are stood in for.
export async function runChoreOver({
  sheetsConfig,
  transport,
  argv,
  now,
}: ChoreRunProps): Promise<ChoreRun> {
  const runner = ChoreRunner.init(sheetsConfig, argv, {
    transport,
    configs: appConfigs,
    now,
  });
  return capturedRun(() => runner.run());
}

interface RestoreRunProps {
  sheetsConfig: SheetsConfig;
  transport: SheetsHttpTransport;
  argv: string[];
}

// Runs a restore past its live transport.
export async function runRestoreOver({
  sheetsConfig,
  transport,
  argv,
}: RestoreRunProps): Promise<ChoreRun> {
  const runner = RestoreRunner.init(sheetsConfig, argv, { transport });
  return capturedRun(() => runner.run());
}

interface ConfirmedSendProps {
  sheetsConfig: SheetsConfig;
  transport: SheetsHttpTransport;
  chore: string;
}

export interface ConfirmedSend {
  dryRun: ChoreRun;
  sent: ChoreRun;
  restore: (argv?: string[]) => Promise<ChoreRun>;
}

// A dry run for the token, the confirmed send it covers, then the restore its undo line names.
export async function confirmedSendOver({
  sheetsConfig,
  transport,
  chore,
}: ConfirmedSendProps): Promise<ConfirmedSend> {
  vi.stubEnv("INIT_CWD", mkdtempSync(join(tmpdir(), "snapshots-")));
  const dryRun = await runChoreOver({ sheetsConfig, transport, argv: [chore] });
  const token = /Token (\S+):/.exec(dryRun.output)?.[1];
  const sent = await runChoreOver({
    sheetsConfig,
    transport,
    argv: [chore, "--send", "--confirm", `${token}`],
  });
  const snapshot = /undo: npm run \S+ -- (\S+)/.exec(sent.output)?.[1];
  return {
    dryRun,
    sent,
    restore: (argv = []) =>
      runRestoreOver({
        sheetsConfig,
        transport,
        argv: [`${snapshot}`, ...argv],
      }),
  };
}

async function capturedRun(run: () => Promise<void>): Promise<ChoreRun> {
  const log = vi.spyOn(console, "log").mockImplementation(() => undefined);
  process.exitCode = undefined;
  await run();
  const output = log.mock.calls.map(([line]) => String(line)).join("\n");
  log.mockRestore();
  const { exitCode } = process;
  // A held or stopped run's exit code would otherwise fail the test process.
  process.exitCode = undefined;
  return { output, exitCode };
}

const srcUrl = new URL("../src/", import.meta.url);

// A chore whose one flush is a findReplace, which the gate always holds.
const replaceRentChoreSource = `
import { googleRawRequest } from ${JSON.stringify(fileURLToPath(new URL("00_Source/GoogleSheets/GoogleSheetsAPI.ts", srcUrl)))};
import { SpreadsheetRaw } from ${JSON.stringify(fileURLToPath(new URL("02_SpreadsheetRaw/SpreadsheetRaw.ts", srcUrl)))};

export const replaceRent = {
  description: "replaces Rent with Lease everywhere",
  action() {
    SpreadsheetRaw.init()
      .gatherRawOperation(
        googleRawRequest({
          findReplace: { find: "Rent", replacement: "Lease", allSheets: true },
        }),
      )
      .batchUpdateGSheets();
  },
};
`;

export function replaceRentSheetsConfig(): SheetsConfig {
  return choreSheetsConfig({
    name: "replaceRent",
    source: replaceRentChoreSource,
  });
}

// A transport that accepts every write and answers every read with an empty spreadsheet.
export function emptySheetsTransport(): SheetsHttpTransport {
  return (request) => {
    if (request.url.includes(":batchUpdate")) return { replies: [] };
    return { sheets: [] };
  };
}
