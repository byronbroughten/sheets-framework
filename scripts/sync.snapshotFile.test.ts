// The scripts project installs no configs, and syncCoordinatorTestSupport reads them as it loads.
import "../dev/installDevConfigs.ts";

import { existsSync, mkdirSync, mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { afterEach, describe, expect, it, vi } from "vitest";

import { appConfigs } from "../dev/generated/appConfigs.ts";
import type { SheetsHttpTransport } from "../src/00_Source/GoogleSheets/GoogleSheetsAPI.ts";
import type { PendingAlterations } from "../src/01_SpreadsheetSchema/configReaders/intendedState.ts";
import { idNameMapFileSource } from "../src/05_Operators/configFileSource.ts";
import {
  bodyRowsWithNote,
  itemColumns,
  itemGid,
  itemRowId,
  previousItemMap,
  seedFixture,
} from "../src/05_Operators/syncCoordinatorTestSupport.ts";
import { fakeSheetsHttpTransport } from "../src/testSupport/fakeSheetsHttpTransport.ts";
import type { FakeSheetsService } from "../src/testSupport/fakeSheetsService.ts";
import type { FakeBodyRow } from "../src/testSupport/fakeSheetsService/fakeTableSheet.ts";
import {
  ageFile,
  type CapturedRun,
  capturedRun,
  clock,
  readJson,
  snapshotFiles,
} from "./runnerFixture.ts";
import type { SheetsConfig } from "./sheetsConfig.ts";
import { SyncRunner } from "./sync.ts";

// The diff shells out to git, and a temporary package is no repo.
vi.mock("./configsDiff.ts", () => ({ readConfigsDiff: () => [] }));

afterEach(() => {
  vi.unstubAllEnvs();
  vi.restoreAllMocks();
});

const noteDeletion: PendingAlterations = {
  tables: { item: { columns: { optionalNote: { deleted: true } } } },
};
const noteRename: PendingAlterations = {
  tables: { item: { columns: { optionalNote: { header: "Mine" } } } },
};
const forceNoteHeader = ["--force", "item/optionalNote/header"];

// A temporary package as an earlier sync left it: the previous ID-name map, and the alterations pending.
function syncSheetsConfig(alterations: PendingAlterations): SheetsConfig {
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

function itemWithNote(): FakeSheetsService {
  return seedFixture({ item: { bodyRows: bodyRowsWithNote } });
}

// A person renamed the note column on the sheet after the last sync.
function itemWithRenamedNote(
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

function itemValues(fake: FakeSheetsService): unknown[] {
  return fake.grid.sheet(itemGid).values().flat();
}

interface SyncRunProps {
  sheetsConfig: SheetsConfig;
  transport: SheetsHttpTransport;
  argv: string[];
}

// A sync past its live host: only the Sheets transport, the clock and the package's tsc are stood in for.
function syncRunner({
  sheetsConfig,
  transport,
  argv,
}: SyncRunProps): SyncRunner {
  const runner = SyncRunner.init(sheetsConfig, argv, {
    transport,
    configs: appConfigs,
    now: () => clock,
  });
  vi.spyOn(runner, "_runTsc").mockReturnValue(true);
  return runner;
}

function runSyncOver(props: SyncRunProps): Promise<CapturedRun> {
  const runner = syncRunner(props);
  return capturedRun(() => runner.run());
}

function snapshotFolder(): string {
  const folder = mkdtempSync(join(tmpdir(), "snapshots-"));
  vi.stubEnv("INIT_CWD", folder);
  return folder;
}

function namedFolderArgs(folder: string): string[] {
  return ["--snapshots", folder, "--restore-script", "app:restore"];
}

// Runs the dry run for its token, and returns the argv of the send that token covers.
async function confirmedArgv(
  props: Omit<SyncRunProps, "argv">,
  extra: string[] = [],
): Promise<string[]> {
  const { output } = await runSyncOver({ ...props, argv: extra });
  const token = /Token (\S+):/.exec(output)?.[1];
  expect(token, output).toBeDefined();
  return ["--send", "--confirm", `${token}`, ...extra];
}

// The folder's one snapshot file, by name and contents.
function onlySnapshot(folder: string) {
  const [name, ...others] = snapshotFiles(folder);
  expect(others).toEqual([]);
  expect(name).toBeDefined();
  return { name: `${name}`, file: readJson(join(folder, `${name}`)) };
}

describe("a sync's confirmed send's snapshot file", () => {
  it("lands in the folder the caller names, with the printed plan, before the undo line naming the caller's restore", async () => {
    const folder = snapshotFolder();
    const fake = itemWithNote();
    const props = {
      sheetsConfig: syncSheetsConfig(noteDeletion),
      transport: fakeSheetsHttpTransport(fake),
    };
    const argv = await confirmedArgv(props);

    const { output, exitCode } = await runSyncOver({
      ...props,
      argv: [...argv, ...namedFolderArgs(folder)],
    });

    expect(exitCode).toBe(0);
    const { name, file } = onlySnapshot(folder);
    expect(file).toMatchObject({
      command: "sync",
      target: expect.stringContaining("sync-send-"),
      takenAt: clock.toISOString(),
    });
    expect(JSON.stringify(file.snapshot)).toContain("kept");
    const plan = `${file.plan}`;
    expect(plan).toContain("item › optionalNote › deleted: delete the column");
    expect(output).toContain(plan);
    expect(
      output.indexOf(`undo: npm run app:restore -- ${name}`),
    ).toBeGreaterThan(output.indexOf(plan));
    expect(itemValues(fake)).not.toContain("kept");
  });

  it("is not written by a dry run, which prints no undo line", async () => {
    const folder = snapshotFolder();
    const fake = itemWithNote();

    const { output } = await runSyncOver({
      sheetsConfig: syncSheetsConfig(noteDeletion),
      transport: fakeSheetsHttpTransport(fake),
      argv: namedFolderArgs(folder),
    });

    expect(snapshotFiles(folder)).toEqual([]);
    expect(output).not.toContain("undo:");
    expect(itemValues(fake)).toContain("kept");
  });

  it("is not written for a send stopped before any flush landed", async () => {
    const folder = snapshotFolder();
    const fake = itemWithNote();
    const answer = fakeSheetsHttpTransport(fake);
    const props = {
      sheetsConfig: syncSheetsConfig(noteDeletion),
      transport: answer,
    };
    const argv = await confirmedArgv(props);

    const { output } = await runSyncOver({
      ...props,
      transport: (request) => {
        if (request.url.includes(":batchUpdate")) {
          throw new Error(
            "Invalid requests[0].deleteRange: You cannot insert or delete cells over part of a table.",
          );
        }
        return answer(request);
      },
      argv: [...argv, ...namedFolderArgs(folder)],
    });

    expect(output).toContain("before any flush landed");
    expect(snapshotFiles(folder)).toEqual([]);
    expect(output).not.toContain("undo:");
  });
});

describe("a sync's forced send's snapshot file", () => {
  it("is written without --confirm, naming the forced trait as undone only by version history", async () => {
    const folder = snapshotFolder();
    const fake = itemWithRenamedNote();

    const { output } = await runSyncOver({
      sheetsConfig: syncSheetsConfig(noteRename),
      transport: fakeSheetsHttpTransport(fake),
      argv: ["--send", ...forceNoteHeader, ...namedFolderArgs(folder)],
    });

    const { name, file } = onlySnapshot(folder);
    expect(output).toContain(`${file.plan}`);
    const afterUndo = output.slice(
      output.indexOf(`undo: npm run app:restore -- ${name}`),
    );
    expect(afterUndo).toContain(
      "undone only by the spreadsheet's version history:",
    );
    expect(afterUndo).toContain("sent over the live value by --force");
    expect(itemValues(fake)).toContain("Mine");
  });

  it("is written, with the undo line, for a send a held flush stopped partway", async () => {
    const folder = snapshotFolder();
    const sheetsConfig = syncSheetsConfig({
      tables: {
        item: {
          columns: {
            optionalNote: { header: "Mine" },
            requiredCount: { deleted: true },
          },
        },
      },
    });
    const argv = await confirmedArgv(
      {
        sheetsConfig,
        transport: fakeSheetsHttpTransport(itemWithRenamedNote()),
      },
      forceNoteHeader,
    );
    // A row a person added after the dry run, which the token doesn't cover.
    const changed = itemWithRenamedNote([
      { id: itemRowId, requiredCount: 1 },
      { id: "r:itm:second", requiredCount: 2 },
    ]);

    const { output, exitCode } = await runSyncOver({
      sheetsConfig,
      transport: fakeSheetsHttpTransport(changed),
      argv: [...argv, ...namedFolderArgs(folder)],
    });

    expect(exitCode).toBe(1);
    expect(output).toContain("STOPPED PARTWAY");
    const { name, file } = onlySnapshot(folder);
    expect(file.plan).toContain("STOPPED PARTWAY");
    expect(output).toContain(`undo: npm run app:restore -- ${name}`);
    expect(itemValues(changed)).toContain("Mine");
    expect(itemValues(changed)).toContain(2);
  });
});

describe("a sync's held send that throws", () => {
  it("still writes the snapshot, with the requests as its plan, and prints the undo line", async () => {
    const folder = snapshotFolder();
    const fake = itemWithNote();
    const answer = fakeSheetsHttpTransport(fake);
    const props = {
      sheetsConfig: syncSheetsConfig(noteDeletion),
      transport: answer,
    };
    const argv = await confirmedArgv(props);
    let hasSent = false;
    const runner = syncRunner({
      ...props,
      transport: (request) => {
        if (hasSent && request.url.includes(":getByDataFilter")) {
          throw new Error("The read after the send failed.");
        }
        if (request.url.includes(":batchUpdate")) hasSent = true;
        return answer(request);
      },
      argv: [...argv, ...namedFolderArgs(folder)],
    });

    const { output } = await capturedRun(() =>
      expect(runner.run()).rejects.toThrow("The read after the send failed."),
    );

    const { name, file } = onlySnapshot(folder);
    expect(file.plan).toContain("delete 1 columns");
    expect(output).toContain(`${file.plan}`);
    expect(output).toContain(`undo: npm run app:restore -- ${name}`);
  });
});

describe("a sync's held send prunes its snapshot folder", () => {
  it("removes snapshots older than 30 days by their file time and keeps newer ones and other files", async () => {
    const folder = snapshotFolder();
    const seeded = {
      old: join(folder, "snapshot-old.json"),
      recent: join(folder, "snapshot-recent.json"),
      other: join(folder, "notes.txt"),
    };
    for (const path of Object.values(seeded)) writeFileSync(path, "{}");
    ageFile(seeded.old, 31);
    ageFile(seeded.recent, 29);
    ageFile(seeded.other, 90);
    const props = {
      sheetsConfig: syncSheetsConfig(noteDeletion),
      transport: fakeSheetsHttpTransport(itemWithNote()),
    };
    const argv = await confirmedArgv(props);

    await runSyncOver({ ...props, argv: [...argv, "--snapshots", folder] });

    expect(existsSync(seeded.old)).toBe(false);
    expect(existsSync(seeded.recent)).toBe(true);
    expect(existsSync(seeded.other)).toBe(true);
    expect(snapshotFiles(folder)).toHaveLength(2);
  });
});
