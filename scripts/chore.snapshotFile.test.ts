import { existsSync } from "node:fs";
import { join } from "node:path";

import { afterEach, describe, expect, it, vi } from "vitest";

import {
  emptySheetsTransport,
  replaceRentSheetsConfig,
  runChoreOver,
} from "./choreFixture.ts";
import { tokenOf } from "./runnerFixture.ts";
import {
  clock,
  readJson,
  seedAgedFiles,
  snapshotFiles,
  snapshotsFolder,
} from "./snapshotFixture.ts";

afterEach(() => {
  vi.unstubAllEnvs();
});

// Runs replaceRent as a dry run, then as the send its token covers.
async function confirmedSend(sheetsConfig = replaceRentSheetsConfig()) {
  function run(argv: string[]) {
    return runChoreOver({
      sheetsConfig,
      transport: emptySheetsTransport(),
      argv: ["replaceRent", ...argv],
      now: () => clock,
    });
  }
  const dryRun = await run([]);
  const token = tokenOf(dryRun.output);
  function sent(extra: string[]) {
    return run(["--send", "--confirm", `${token}`, ...extra]);
  }
  return { dryRun, sent, sheetsConfig };
}

describe("a chore's confirmed send's snapshot file", () => {
  it("lands in the folder the caller names, with the plan, and the undo line names the restore the caller names", async () => {
    const folder = snapshotsFolder();
    const { sent } = await confirmedSend();

    const { output } = await sent([
      "--snapshots",
      folder,
      "--restore-script",
      "app:restore",
    ]);

    const [name, ...others] = snapshotFiles(folder);
    expect(others).toEqual([]);
    expect(output).toContain(`undo: npm run app:restore -- ${name}`);
    const file = readJson(join(folder, `${name}`));
    expect(file).toMatchObject({
      command: "chore replaceRent",
      takenAt: clock.toISOString(),
      snapshot: { columns: [], fingerprints: [] },
    });
    const plan = `${file.plan}`;
    expect(plan).toContain("findReplace");
    expect(output).toContain(plan);
  });

  it("names the held items the snapshot can't cover beside the undo line", async () => {
    const folder = snapshotsFolder();
    const { sent } = await confirmedSend();

    const { output } = await sent(["--snapshots", folder]);

    const afterUndo = output.slice(output.indexOf("undo: "));
    expect(afterUndo).toContain(
      "undone only by the spreadsheet's version history:",
    );
    expect(afterUndo).toContain(
      'a find-and-replace of "Rent" with "Lease", which can reach entered data anywhere',
    );
  });

  it("falls back to the package's .probe/snapshots and its own restore script", async () => {
    const { sent, sheetsConfig } = await confirmedSend();
    vi.stubEnv("INIT_CWD", sheetsConfig.dir);

    const { output } = await sent([]);

    const folder = join(sheetsConfig.dir, ".probe", "snapshots");
    const [name] = snapshotFiles(folder);
    expect(snapshotFiles(folder)).toHaveLength(1);
    expect(output).toContain(
      `undo: npm run restore -- ${join(".probe", "snapshots", `${name}`)}`,
    );
  });

  it("is not written by a dry run, which prints no undo line", async () => {
    const folder = snapshotsFolder();
    const sheetsConfig = replaceRentSheetsConfig();

    const dryRun = await runChoreOver({
      sheetsConfig,
      transport: emptySheetsTransport(),
      argv: ["replaceRent", "--snapshots", folder, "--restore-script", "x"],
      now: () => clock,
    });

    expect(snapshotFiles(folder)).toEqual([]);
    expect(dryRun.output).not.toContain("undo:");
  });
});

describe("a chore's confirmed send prunes its snapshot folder", () => {
  it("removes snapshots older than 30 days by their file time and keeps newer ones and other files", async () => {
    const folder = snapshotsFolder();
    const seeded = seedAgedFiles(folder);
    const { sent } = await confirmedSend();

    await sent(["--snapshots", folder]);

    expect(existsSync(seeded.old)).toBe(false);
    expect(existsSync(seeded.recent)).toBe(true);
    expect(existsSync(seeded.other)).toBe(true);
    expect(snapshotFiles(folder)).toHaveLength(2);
  });
});
