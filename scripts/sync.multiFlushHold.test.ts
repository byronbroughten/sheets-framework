// The scripts project installs no configs, and syncCoordinatorTestSupport reads them as it loads.
import "../dev/installDevConfigs.ts";

import { join } from "node:path";

import { afterEach, describe, expect, it, vi } from "vitest";

import type { PendingAlterations } from "../src/01_SpreadsheetSchema/configReaders/intendedState.ts";
import {
  itemRowId,
  seedFixture,
} from "../src/05_Operators/syncCoordinatorTestSupport.ts";
import { fakeSheetsHttpTransport } from "../src/testSupport/fakeSheetsHttpTransport.ts";
import type { FakeSheetsService } from "../src/testSupport/fakeSheetsService.ts";
import { readJson, snapshotFiles, snapshotsFolder } from "./snapshotFixture.ts";
import {
  confirmedArgv,
  forceNoteHeader,
  itemValues,
  itemWithRenamedNote,
  runSyncOver,
  syncSheetsConfig,
} from "./syncFixture.ts";

// The diff shells out to git, and a temporary package is no repo.
vi.mock("./configsDiff.ts", () => ({ readConfigsDiff: () => [] }));

afterEach(() => {
  vi.unstubAllEnvs();
  vi.restoreAllMocks();
});

const heldCount = `r:itm:fixture  "1"  → deleted`;
const countDeletion = { requiredCount: { deleted: true } } as const;
const noteRenameWithCountDeletion: PendingAlterations = {
  tables: {
    item: { columns: { optionalNote: { header: "Mine" }, ...countDeletion } },
  },
};

interface HoldCase {
  alterations: PendingAlterations;
  // The item Table as the dry run reads it, given the row's count.
  seed: (count: number) => FakeSheetsService;
  extra: string[];
}

// The note column renamed on the sheet, so only --force sends the alteration's header over it.
const forcedPair: HoldCase = {
  alterations: noteRenameWithCountDeletion,
  seed: (count) =>
    itemWithRenamedNote([{ id: itemRowId, requiredCount: count }]),
  extra: forceNoteHeader,
};
// The note column as the last sync left it, so the rename sends without --force.
const plainPair: HoldCase = {
  alterations: noteRenameWithCountDeletion,
  seed: (count) =>
    seedFixture({
      item: { bodyRows: [{ id: itemRowId, requiredCount: count }] },
    }),
  extra: [],
};
const deletionAlone: HoldCase = {
  alterations: { tables: { item: { columns: countDeletion } } },
  seed: forcedPair.seed,
  extra: [],
};

async function dryRunOutput({ alterations, seed, extra }: HoldCase) {
  const { output } = await runSyncOver({
    sheetsConfig: syncSheetsConfig(alterations),
    transport: fakeSheetsHttpTransport(seed(1)),
    argv: extra,
  });
  return output;
}

// The dry run over a count of 1, then its confirmed send over a count of sendCount.
async function confirmedSend(
  { alterations, seed, extra }: HoldCase,
  sendCount: number,
  folderArgs: string[] = [],
) {
  const sheetsConfig = syncSheetsConfig(alterations);
  const argv = await confirmedArgv(
    { sheetsConfig, transport: fakeSheetsHttpTransport(seed(1)) },
    extra,
  );
  const fake = seed(sendCount);
  const run = await runSyncOver({
    sheetsConfig,
    transport: fakeSheetsHttpTransport(fake),
    argv: [...argv, ...folderArgs],
  });
  return { ...run, fake };
}

describe("a sync whose deletion flush follows a non-empty flush", () => {
  for (const [label, holdCase] of [
    ["a forced header", forcedPair],
    ["a plain header rename", plainPair],
  ] as const) {
    describe(`after ${label}`, () => {
      it("holds the deleted column's entered value in the dry run", async () => {
        expect(await dryRunOutput(holdCase)).toContain(heldCount);
      });

      it("stops a send whose value changed since the dry run", async () => {
        const { output, exitCode, fake } = await confirmedSend(holdCase, 2);

        expect(exitCode, output).toBe(1);
        expect(output).toContain("STOPPED PARTWAY");
        expect(itemValues(fake)).toContain(2);
      });

      it("lands a send whose value is unchanged", async () => {
        const { output, exitCode, fake } = await confirmedSend(holdCase, 1);

        expect(exitCode, output).toBe(0);
        expect(itemValues(fake)).toContain("Mine");
        expect(itemValues(fake)).not.toContain(1);
      });
    });
  }

  it("snapshots the deleted column's entered values on a forced send", async () => {
    const folder = snapshotsFolder();

    await confirmedSend(forcedPair, 1, ["--snapshots", folder]);

    const [name] = snapshotFiles(folder);
    const file = readJson(join(folder, `${name}`));
    expect(file.snapshot).toMatchObject({
      columns: [
        expect.objectContaining({
          columnId: "c:itm:requiredCount",
          values: expect.arrayContaining([1]),
        }),
      ],
    });
  });

  it("still leaves the forced trait itself out of the hold", async () => {
    const output = await dryRunOutput({
      ...forcedPair,
      alterations: {
        tables: { item: { columns: { optionalNote: { header: "Mine" } } } },
      },
    });

    expect(output).not.toContain("== Held");
    expect(output).not.toContain("Token ");
  });
});

describe("a sync whose only flush deletes a column", () => {
  it("holds the deleted column's entered value in the dry run", async () => {
    expect(await dryRunOutput(deletionAlone)).toContain(heldCount);
  });

  it("stops a send whose value changed since the dry run", async () => {
    const { output, exitCode, fake } = await confirmedSend(deletionAlone, 2);

    expect(exitCode, output).toBe(1);
    expect(itemValues(fake)).toContain(2);
  });
});
