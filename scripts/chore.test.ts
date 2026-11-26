import { describe, expect, it, vi } from "vitest";

import type { SheetsHttpRequest } from "../src/00_Source/GoogleSheets/GoogleSheetsAPI.ts";
import {
  type ChoreRun,
  replaceRentSheetsConfig,
  runChoreOver,
} from "./choreFixture.ts";

const sheetsConfig = replaceRentSheetsConfig();

async function runReplaceRent(
  argv: string[],
): Promise<ChoreRun & { batchUpdates: number }> {
  const transport = vi.fn((request: SheetsHttpRequest) => {
    if (request.url.includes(":batchUpdate")) return { replies: [] };
    return { sheets: [] };
  });
  const run = await runChoreOver({
    sheetsConfig,
    transport,
    argv: ["replaceRent", ...argv],
  });
  return {
    ...run,
    batchUpdates: transport.mock.calls.filter(([request]) =>
      request.url.includes(":batchUpdate"),
    ).length,
  };
}

describe("a chore's --send over entered data", () => {
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
    expect(dryRun.output).not.toContain("Re-run with `-- --send` to apply.");
    expect(dryRun.batchUpdates).toBe(0);

    const sent = await runReplaceRent(["--send", "--confirm", `${token}`]);
    expect(sent.batchUpdates).toBe(1);
    expect(sent.exitCode).toBeUndefined();
    expect(sent.output).toContain("SENT — the spreadsheet was written to:");
  });
});
