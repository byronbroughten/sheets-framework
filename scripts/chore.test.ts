import { createHash } from "node:crypto";
import { mkdirSync, mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

import { afterEach, describe, expect, it, vi } from "vitest";

import { appConfigs } from "../dev/generated/appConfigs.ts";
import type { SheetsHttpRequest } from "../src/00_Source/GoogleSheets/GoogleSheetsAPI.ts";
import { NodeHost } from "../src/nodeHost/NodeHost.ts";
import { ChoreRunner } from "./chore.ts";
import { heldFile } from "./heldFile.ts";
import type { SheetsConfig } from "./sheetsConfig.ts";

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

const packageDir = mkdtempSync(join(tmpdir(), "chore-"));
const choreHome = join(packageDir, "chores");
mkdirSync(choreHome);
writeFileSync(join(choreHome, "replaceRent.ts"), replaceRentChoreSource);

const sheetsConfig: SheetsConfig = {
  path: join(packageDir, "sheets.config.json"),
  dir: packageDir,
  spreadsheetId: "spreadsheet-under-test",
  generatedDir: join(packageDir, "generated"),
  choreHomes: [choreHome],
  endpointModule: undefined,
};

interface ChoreRun {
  output: string;
  exitCode: number | string | null | undefined;
  batchUpdates: number;
}

// Runs the chore past its live host: only the Sheets transport is stood in for.
async function runReplaceRent(argv: string[]): Promise<ChoreRun> {
  const runner = ChoreRunner.init(sheetsConfig, ["replaceRent", ...argv]);
  const transport = vi.fn((request: SheetsHttpRequest) => {
    if (request.url.includes(":batchUpdate")) return { replies: [] };
    return { sheets: [] };
  });
  vi.spyOn(runner, "_startHost").mockImplementation(async () =>
    NodeHost.init({
      configs: appConfigs,
      spreadsheetId: sheetsConfig.spreadsheetId,
      transport,
      isDryRun: !runner.isSend,
      log: vi.fn(),
      enteredDataGate: {
        target: runner.target,
        confirmedList: runner._confirmedList(),
        hash: (input) => createHash("sha256").update(input).digest("hex"),
      },
    }).ensureGlobals(),
  );
  const log = vi.spyOn(console, "log").mockImplementation(() => undefined);
  process.exitCode = undefined;
  await runner.run();
  const output = log.mock.calls.map(([line]) => String(line)).join("\n");
  log.mockRestore();
  return {
    output,
    exitCode: process.exitCode,
    batchUpdates: transport.mock.calls.filter(([request]) =>
      request.url.includes(":batchUpdate"),
    ).length,
  };
}

describe("a chore's --send over entered data", () => {
  afterEach(() => {
    process.exitCode = undefined;
    vi.restoreAllMocks();
  });

  it("keeps its held list in its own held file", () => {
    expect(heldFile.path(packageDir, "chore")).toBe(
      join(packageDir, ".probe", "chore-held.json"),
    );
  });

  it("is held without --confirm, then sent once --confirm gives the dry run's token", async () => {
    const refused = await runReplaceRent(["--send"]);
    expect(refused.batchUpdates).toBe(0);
    expect(refused.exitCode).toBe(1);
    expect(refused.output).toContain(
      "STOPPED: the send was refused before any flush landed",
    );
    expect(refused.output).toContain(
      "Rerun the dry run: npm run chore replaceRent",
    );

    const dryRun = await runReplaceRent([]);
    const token = /Token (\S+):/.exec(dryRun.output)?.[1];
    expect(dryRun.output).toContain(
      "== Held: 1 item, sent only with --confirm ==",
    );
    expect(dryRun.output).toContain(
      `Rerun with --send --confirm ${token} to send these.`,
    );
    expect(dryRun.batchUpdates).toBe(0);

    const sent = await runReplaceRent(["--send", "--confirm", `${token}`]);
    expect(sent.batchUpdates).toBe(1);
    expect(sent.exitCode).toBeUndefined();
    expect(sent.output).toContain("SENT — the spreadsheet was written to:");
  });
});
