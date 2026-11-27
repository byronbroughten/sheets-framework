// The scripts project installs no configs, and syncCoordinatorTestSupport reads them as it loads.
import "../dev/installDevConfigs.ts";

import { existsSync, mkdirSync, rmSync } from "node:fs";
import { join } from "node:path";

import { afterEach, describe, expect, it, vi } from "vitest";

import type { SheetsHttpTransport } from "../src/00_Source/GoogleSheets/GoogleSheetsAPI.ts";
import type { PendingAlterations } from "../src/01_SpreadsheetSchema/configReaders/intendedState.ts";
import {
  bodyRowsWithNote,
  itemRowId,
  seedFixture,
} from "../src/05_Operators/syncCoordinatorTestSupport.ts";
import { fakeSheetsHttpTransport } from "../src/testSupport/fakeSheetsHttpTransport.ts";
import type { FakeSheetsService } from "../src/testSupport/fakeSheetsService.ts";
import { capturedRun } from "./runnerFixture.ts";
import {
  clock,
  onlySnapshot,
  seedAgedFiles,
  snapshotFiles,
  snapshotsFolder,
} from "./snapshotFixture.ts";
import {
  confirmedArgv,
  forceNoteHeader,
  itemValues,
  itemWithRenamedNote,
  runSyncOver,
  syncRunner,
  syncSheetsConfig,
} from "./syncFixture.ts";

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
function itemWithNote(): FakeSheetsService {
  return seedFixture({ item: { bodyRows: bodyRowsWithNote } });
}

function namedFolderArgs(folder: string): string[] {
  return ["--snapshots", folder, "--restore-script", "app:restore"];
}

// Throws each failure on its matching requests once a flush has gone out, and answers every other one.
function failingAfterSend(
  answer: SheetsHttpTransport,
  failures: Record<string, string>,
): SheetsHttpTransport {
  let hasSent = false;
  return (request) => {
    for (const [urlPart, message] of Object.entries(failures)) {
      if (hasSent && request.url.includes(urlPart)) throw new Error(message);
    }
    if (request.url.includes(":batchUpdate")) hasSent = true;
    return answer(request);
  };
}

describe("a sync's confirmed send's snapshot file", () => {
  it("lands in the folder the caller names, with the printed plan, before the undo line naming the caller's restore", async () => {
    const folder = snapshotsFolder();
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
    const folder = snapshotsFolder();
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
    const folder = snapshotsFolder();
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
    const folder = snapshotsFolder();
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
    const folder = snapshotsFolder();
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
    const folder = snapshotsFolder();
    const fake = itemWithNote();
    const answer = fakeSheetsHttpTransport(fake);
    const props = {
      sheetsConfig: syncSheetsConfig(noteDeletion),
      transport: answer,
    };
    const argv = await confirmedArgv(props);
    const runner = syncRunner({
      ...props,
      transport: failingAfterSend(answer, {
        ":getByDataFilter": "The read after the send failed.",
      }),
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

describe("a sync's held send whose snapshot or generated files fail", () => {
  it("says only version history undoes a send whose snapshot is lost, still writes the generated files, and exits 1", async () => {
    const folder = snapshotsFolder();
    const sheetsConfig = syncSheetsConfig(noteDeletion);
    const answer = fakeSheetsHttpTransport(itemWithNote());
    const argv = await confirmedArgv({ sheetsConfig, transport: answer });
    // The dry run wrote it too, so only the send can put it back.
    const tableConfigs = join(sheetsConfig.generatedDir, "tableConfigs.ts");
    rmSync(tableConfigs);

    const { output, exitCode } = await runSyncOver({
      sheetsConfig,
      transport: failingAfterSend(answer, {
        "/values:batchGet": "The snapshot read failed.",
      }),
      argv: [...argv, ...namedFolderArgs(folder)],
    });

    expect(exitCode).toBe(1);
    expect(output).toContain(
      "sync: no snapshot was written, so only the spreadsheet's version history undoes this send: The snapshot read failed.",
    );
    expect(existsSync(tableConfigs)).toBe(true);
    expect(snapshotFiles(folder)).toEqual([]);
    expect(output).not.toContain("undo:");
  });

  it("keeps a thrown send's own error when the snapshot is lost too", async () => {
    const folder = snapshotsFolder();
    const sheetsConfig = syncSheetsConfig(noteDeletion);
    const answer = fakeSheetsHttpTransport(itemWithNote());
    const argv = await confirmedArgv({ sheetsConfig, transport: answer });
    const runner = syncRunner({
      sheetsConfig,
      transport: failingAfterSend(answer, {
        ":getByDataFilter": "The read after the send failed.",
        "/values:batchGet": "The snapshot read failed.",
      }),
      argv: [...argv, ...namedFolderArgs(folder)],
    });

    const { output } = await capturedRun(() =>
      expect(runner.run()).rejects.toThrow("The read after the send failed."),
    );

    expect(output).toContain(
      "version history undoes this send: The snapshot read failed.",
    );
    expect(snapshotFiles(folder)).toEqual([]);
  });

  it("still prints the undo line when a generated file can't be written", async () => {
    const folder = snapshotsFolder();
    const sheetsConfig = syncSheetsConfig(noteDeletion);
    const transport = fakeSheetsHttpTransport(itemWithNote());
    const argv = await confirmedArgv({ sheetsConfig, transport });
    // A folder where the file goes, so writing it throws.
    const tableConfigs = join(sheetsConfig.generatedDir, "tableConfigs.ts");
    rmSync(tableConfigs);
    mkdirSync(tableConfigs);
    const runner = syncRunner({
      sheetsConfig,
      transport,
      argv: [...argv, ...namedFolderArgs(folder)],
    });

    const { output } = await capturedRun(() =>
      expect(runner.run()).rejects.toThrow("EISDIR"),
    );

    const { name } = onlySnapshot(folder);
    expect(output).toContain(`undo: npm run app:restore -- ${name}`);
  });
});

describe("a sync's held send prunes its snapshot folder", () => {
  it("removes snapshots older than 30 days by their file time and keeps newer ones and other files", async () => {
    const folder = snapshotsFolder();
    const seeded = seedAgedFiles(folder);
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
