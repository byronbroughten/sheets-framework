import { mkdirSync, mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

import { expect } from "vitest";

import { appConfigs } from "../dev/generated/appConfigs.ts";
import type { SheetsHttpTransport } from "../src/00_Source/GoogleSheets/GoogleSheetsAPI.ts";
import { ChoreRunner } from "./chore.ts";
import { RestoreRunner } from "./restore.ts";
import { type CapturedRun, capturedRun, tokenOf } from "./runnerFixture.ts";
import type { SheetsConfig } from "./sheetsConfig.ts";
import { snapshotsFolder } from "./snapshotFixture.ts";

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
}: ChoreRunProps): Promise<CapturedRun> {
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
}: RestoreRunProps): Promise<CapturedRun> {
  const runner = RestoreRunner.init(sheetsConfig, argv, { transport });
  return capturedRun(() => runner.run());
}

interface ConfirmedSendProps {
  sheetsConfig: SheetsConfig;
  transport: SheetsHttpTransport;
  chore: string;
}

export interface ConfirmedSend {
  dryRun: CapturedRun;
  sent: CapturedRun;
  restore: (argv?: string[]) => Promise<CapturedRun>;
}

// A dry run for the token, the confirmed send it covers, which must print an undo line, then the restore that line names.
export async function confirmedSendOver({
  sheetsConfig,
  transport,
  chore,
}: ConfirmedSendProps): Promise<ConfirmedSend> {
  snapshotsFolder();
  const dryRun = await runChoreOver({ sheetsConfig, transport, argv: [chore] });
  const token = tokenOf(dryRun.output);
  const sent = await runChoreOver({
    sheetsConfig,
    transport,
    argv: [chore, "--send", "--confirm", `${token}`],
  });
  expect(sent.output, `${dryRun.output}\n---\n${sent.output}`).toMatch(
    /undo: /,
  );
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

const srcUrl = new URL("../src/", import.meta.url);

// A src/ module's path as a chore source's import specifier.
export function srcModule(path: string): string {
  return JSON.stringify(fileURLToPath(new URL(path, srcUrl)));
}

// A chore whose one flush is a findReplace, which the gate always holds.
const replaceRentChoreSource = `
import { googleRawRequest } from ${srcModule("00_Source/GoogleSheets/GoogleSheetsAPI.ts")};
import { SpreadsheetRaw } from ${srcModule("02_SpreadsheetRaw/SpreadsheetRaw.ts")};

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
