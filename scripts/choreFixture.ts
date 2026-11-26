import { mkdirSync, mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { vi } from "vitest";

import { appConfigs } from "../dev/generated/appConfigs.ts";
import type { SheetsHttpTransport } from "../src/00_Source/GoogleSheets/GoogleSheetsAPI.ts";
import { ChoreRunner } from "./chore.ts";
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
}

// Runs a chore past its live host: only the Sheets transport is stood in for.
export async function runChoreOver({
  sheetsConfig,
  transport,
  argv,
}: ChoreRunProps): Promise<ChoreRun> {
  const runner = ChoreRunner.init(sheetsConfig, argv, {
    transport,
    configs: appConfigs,
  });
  const log = vi.spyOn(console, "log").mockImplementation(() => undefined);
  process.exitCode = undefined;
  await runner.run();
  const output = log.mock.calls.map(([line]) => String(line)).join("\n");
  log.mockRestore();
  const { exitCode } = process;
  // A held run's exit code would otherwise fail the test process.
  process.exitCode = undefined;
  return { output, exitCode };
}
