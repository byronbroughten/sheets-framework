// What the sync runner's tests over the fake share: a temporary package, the item Table a person changed, and a run with its confirmed send.
import { mkdirSync, mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { expect } from "vitest";

import { appConfigs } from "../dev/generated/appConfigs.ts";
import type { SheetsHttpTransport } from "../src/00_Source/GoogleSheets/GoogleSheetsAPI.ts";
import type { PendingAlterations } from "../src/01_SpreadsheetSchema/configReaders/intendedState.ts";
import { idNameMapFileSource } from "../src/05_Operators/configFileSource.ts";
import {
  itemColumns,
  itemRowId,
  previousItemMap,
  seedFixture,
} from "../src/05_Operators/syncCoordinatorTestSupport.ts";
import type { FakeSheetsService } from "../src/testSupport/fakeSheetsService.ts";
import type { FakeBodyRow } from "../src/testSupport/fakeSheetsService/fakeTableSheet.ts";
import { type CapturedRun, capturedRun, tokenOf } from "./runnerFixture.ts";
import type { SheetsConfig } from "./sheetsConfig.ts";
import { clock } from "./snapshotFixture.ts";
import { SyncRunner } from "./sync.ts";

export const forceNoteHeader = ["--force", "item/optionalNote/header"];

// A temporary package as an earlier sync left it: the previous ID-name map, and the alterations pending.
export function syncSheetsConfig(
  alterations: PendingAlterations,
): SheetsConfig {
  const packageDir = mkdtempSync(join(tmpdir(), "sync-send-"));
  const generatedDir = join(packageDir, "generated");
  mkdirSync(generatedDir);
  writeFileSync(
    join(generatedDir, "idNameMap.ts"),
    idNameMapFileSource(previousItemMap()),
  );
  writeFileSync(
    join(packageDir, "alterations.ts"),
    `export const pendingAlterations = ${JSON.stringify(alterations)};\n`,
  );
  return {
    path: join(packageDir, "sheets.config.json"),
    dir: packageDir,
    spreadsheetId: "spreadsheet-under-test",
    generatedDir,
    choreHomes: [],
    endpointModule: undefined,
  };
}

// A person renamed the note column on the sheet after the last sync.
export function itemWithRenamedNote(
  bodyRows: FakeBodyRow<string>[] = [{ id: itemRowId, requiredCount: 1 }],
): FakeSheetsService {
  return seedFixture({
    item: {
      columns: {
        optionalNote: { ...itemColumns.optionalNote, header: "Theirs" },
      },
      bodyRows,
    },
  });
}

export interface SyncRunProps {
  sheetsConfig: SheetsConfig;
  transport: SheetsHttpTransport;
  argv: string[];
}

// A sync past its live host: the transport, the dev configs, the clock and tsc are stood in for; the test mocks the git diff.
export function syncRunner({
  sheetsConfig,
  transport,
  argv,
}: SyncRunProps): SyncRunner {
  return SyncRunner.init(sheetsConfig, argv, {
    transport,
    configs: appConfigs,
    now: () => clock,
    runTsc: () => true,
  });
}

export function runSyncOver(props: SyncRunProps): Promise<CapturedRun> {
  const runner = syncRunner(props);
  return capturedRun(() => runner.run());
}

// Runs the dry run for its token, and returns the argv of the send that token covers.
export async function confirmedArgv(
  props: Omit<SyncRunProps, "argv">,
  extra: string[] = [],
): Promise<string[]> {
  const { output } = await runSyncOver({ ...props, argv: extra });
  const token = tokenOf(output);
  expect(token, output).toBeDefined();
  return ["--send", "--confirm", `${token}`, ...extra];
}
