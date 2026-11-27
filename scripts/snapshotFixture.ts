// What the snapshot file's tests share: the injected clock, a snapshot folder, its files and their ages.
import {
  existsSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  utimesSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { expect, vi } from "vitest";

export const clock = new Date("2026-11-25T12:00:00.000Z");

const dayMilliseconds = 24 * 60 * 60 * 1000;

// A new folder npm counts as invoked from, so an undo line names its file bare; the test unstubs the env.
export function snapshotsFolder(): string {
  const folder = mkdtempSync(join(tmpdir(), "snapshots-"));
  vi.stubEnv("INIT_CWD", folder);
  return folder;
}

export function snapshotFiles(folder: string): string[] {
  if (!existsSync(folder)) return [];
  return readdirSync(folder).filter((name) => name.startsWith("snapshot-"));
}

export function readJson(path: string): Record<string, unknown> {
  return JSON.parse(readFileSync(path, "utf8"));
}

export interface SnapshotFile {
  name: string;
  file: Record<string, unknown>;
}

// The folder's one snapshot file, by name and contents.
export function onlySnapshot(folder: string): SnapshotFile {
  const [name, ...others] = snapshotFiles(folder);
  expect(others).toEqual([]);
  expect(name).toBeDefined();
  return { name: `${name}`, file: readJson(join(folder, `${name}`)) };
}

export interface AgedFiles {
  old: string;
  recent: string;
  other: string;
}

// An old snapshot, a recent one and an older file that is no snapshot, for a prune to meet.
export function seedAgedFiles(folder: string): AgedFiles {
  const seeded = {
    old: join(folder, "snapshot-old.json"),
    recent: join(folder, "snapshot-recent.json"),
    other: join(folder, "notes.txt"),
  };
  for (const path of Object.values(seeded)) writeFileSync(path, "{}");
  ageFile(seeded.old, 31);
  ageFile(seeded.recent, 29);
  ageFile(seeded.other, 90);
  return seeded;
}

// Seeds a file's time this many days before the clock, so a prune meets an old file with no wait.
function ageFile(path: string, daysOld: number): void {
  const time = new Date(clock.getTime() - daysOld * dayMilliseconds);
  utimesSync(path, time, time);
}
