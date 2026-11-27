// The scripts project installs no configs, and valueTypesTableSheet reads them as it loads.
import "../dev/installDevConfigs.ts";

import { join } from "node:path";

import { afterEach, describe, expect, it, vi } from "vitest";

import {
  bodyRowsWithNote,
  itemColumns,
  itemGid,
  seedFixture,
} from "../src/05_Operators/syncCoordinatorTestSupport.ts";
import { fakeSheetsHttpTransport } from "../src/testSupport/fakeSheetsHttpTransport.ts";
import type { FakeSheetsService } from "../src/testSupport/fakeSheetsService.ts";
import {
  type ValueTypesBodyRow,
  valueTypesGid,
  valueTypesTableSheet,
} from "../src/testSupport/valueTypesTableSheet.ts";
import {
  choreSheetsConfig,
  type ConfirmedSend,
  confirmedSendOver,
} from "./choreFixture.ts";
import { colOf, withRowIds } from "./restoreFixture.ts";
import { onlySnapshot } from "./snapshotFixture.ts";

afterEach(() => {
  vi.unstubAllEnvs();
});

const bodyRows: ValueTypesBodyRow[] = [
  { id: "r:vty:apple", stringValue: "Apple", checkbox: true },
  { id: "r:vty:pear", stringValue: "Pear", checkbox: false },
];

const stringValueColIndex = colOf(withRowIds, "stringValue");
const noteColIndex = Object.keys(itemColumns).indexOf("optionalNote");

// Deletes a Value Types column and renames its tab in the same flush.
const renameChoreSource = `
export const renameTab = {
  description: "deletes a Value Types column and renames its tab",
  action(ss) {
    ss.table("valueTypes").prepFetchColumnsFull("stringValue");
    ss.fetchAllPrepped();
    const table = ss.table("valueTypes").raw;
    table.deleteColumn(${stringValueColIndex});
    table.sheet.updateTitle("Renamed Types");
    ss.batchUpdateGSheets();
  },
};
`;

// Deletes a column on each tab, then the Value Types tab in a later flush.
const deleteChoreSource = `
export const deleteTab = {
  description: "deletes a column on each tab, then the Value Types tab",
  action(ss) {
    ss.table("valueTypes").prepFetchColumnsFull("stringValue");
    ss.table("item").prepFetchColumnsFull("optionalNote");
    ss.fetchAllPrepped();
    ss.table("valueTypes").raw.deleteColumn(${stringValueColIndex});
    ss.table("item").raw.deleteColumn(${noteColIndex});
    ss.batchUpdateGSheets();
    ss.table("valueTypes").raw.sheet.queueDelete();
    ss.batchUpdateGSheets();
  },
};
`;

interface TabChangingSend extends ConfirmedSend {
  fake: FakeSheetsService;
  snapshot: Record<string, unknown>;
}

function twoTabs(): FakeSheetsService {
  return seedFixture({
    item: { bodyRows: bodyRowsWithNote },
    extraSheets: [valueTypesTableSheet(bodyRows)],
  });
}

// The confirmed send, and the one snapshot file it wrote to the package's default folder.
async function tabChangingSend(
  fake: FakeSheetsService,
  chore: string,
  source: string,
): Promise<TabChangingSend> {
  const sheetsConfig = choreSheetsConfig({ name: chore, source });
  const sent = await confirmedSendOver({
    sheetsConfig,
    transport: fakeSheetsHttpTransport(fake),
    chore,
  });
  const folder = join(sheetsConfig.dir, ".probe", "snapshots");
  return { ...sent, fake, snapshot: onlySnapshot(folder).file };
}

describe("a chore's confirmed send that renames a tab with captured columns", () => {
  it("writes its snapshot, with the tab's fingerprints read under its new title, and prints the undo line", async () => {
    const fake = twoTabs();

    const { sent, snapshot } = await tabChangingSend(
      fake,
      "renameTab",
      renameChoreSource,
    );

    expect(sent.output).toMatch(/undo: npm run \S+ -- \S+/);
    expect(fake.grid.sheet(valueTypesGid).title).toBe("Renamed Types");
    expect(snapshot).toMatchObject({
      snapshot: {
        fingerprints: [
          {
            sheetGid: valueTypesGid,
            after: expect.arrayContaining([
              expect.arrayContaining(["r:vty:apple"]),
              expect.arrayContaining(["r:vty:pear"]),
            ]),
          },
        ],
      },
    });
  });

  it("restores the renamed tab's column", async () => {
    const fake = twoTabs();
    const gridBefore = fake.grid.sheet(valueTypesGid).values();
    const { restore } = await tabChangingSend(
      fake,
      "renameTab",
      renameChoreSource,
    );

    const { output, exitCode } = await restore();

    expect(exitCode).toBeUndefined();
    expect(output).not.toContain("not restored");
    expect(fake.grid.sheet(valueTypesGid).values()).toEqual(gridBefore);
  });
});

describe("a chore's confirmed send that deletes a tab with captured columns", () => {
  it("writes its snapshot, with the other tab's fingerprints and every saved column, and prints the undo line", async () => {
    const { sent, snapshot } = await tabChangingSend(
      twoTabs(),
      "deleteTab",
      deleteChoreSource,
    );

    expect(sent.output).toMatch(/undo: npm run \S+ -- \S+/);
    expect(snapshot).toMatchObject({
      snapshot: {
        columns: expect.arrayContaining([
          expect.objectContaining({ sheetGid: valueTypesGid }),
          expect.objectContaining({ sheetGid: itemGid }),
        ]),
        fingerprints: expect.arrayContaining([
          expect.objectContaining({
            sheetGid: valueTypesGid,
            isGoneAfterSend: true,
          }),
          expect.objectContaining({
            sheetGid: itemGid,
            after: expect.arrayContaining([
              expect.arrayContaining(["r:itm:second"]),
            ]),
          }),
        ]),
      },
    });
  });

  it("restores the tab that remains and lists the deleted tab's column as not restored", async () => {
    const fake = twoTabs();
    const itemGridBefore = fake.grid.sheet(itemGid).values();
    const { restore } = await tabChangingSend(
      fake,
      "deleteTab",
      deleteChoreSource,
    );
    expect(fake.grid.sheet(itemGid).values()).not.toEqual(itemGridBefore);

    const { output, exitCode } = await restore();

    expect(exitCode).toBeUndefined();
    expect(output).not.toContain("RESTORE STOPPED");
    expect(fake.grid.sheet(itemGid).values()).toEqual(itemGridBefore);
    const notRestored = output.slice(output.indexOf("not restored"));
    expect(notRestored).toContain("Value Types › ");
    expect(notRestored).not.toContain("Item › ");
  });
});
