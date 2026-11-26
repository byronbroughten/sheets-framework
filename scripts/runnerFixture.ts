// What the script runners' tests share: a run's captured output, and a snapshot folder's files and their ages.
import { existsSync, readdirSync, readFileSync, utimesSync } from "node:fs";

import { vi } from "vitest";

export const clock = new Date("2026-11-25T12:00:00.000Z");

const dayMilliseconds = 24 * 60 * 60 * 1000;

export interface CapturedRun {
  output: string;
  exitCode: number | string | null | undefined;
}

export async function capturedRun(
  run: () => Promise<void>,
): Promise<CapturedRun> {
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

export function snapshotFiles(folder: string): string[] {
  if (!existsSync(folder)) return [];
  return readdirSync(folder).filter((name) => name.startsWith("snapshot-"));
}

export function readJson(path: string): Record<string, unknown> {
  return JSON.parse(readFileSync(path, "utf8"));
}

// Seeds a file's time this many days before the clock, so a prune meets an old file with no wait.
export function ageFile(path: string, daysOld: number): void {
  const time = new Date(clock.getTime() - daysOld * dayMilliseconds);
  utimesSync(path, time, time);
}
