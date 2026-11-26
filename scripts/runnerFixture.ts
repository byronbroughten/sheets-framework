// What the script runners' tests share: a run's captured output, and the token its dry run printed.
import { vi } from "vitest";

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

export function tokenOf(output: string): string | undefined {
  return /Token (\S+):/.exec(output)?.[1];
}
