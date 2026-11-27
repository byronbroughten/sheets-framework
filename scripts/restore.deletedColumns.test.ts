// The scripts project installs no configs, and valueTypesTableSheet reads them as it loads.
import "../dev/installDevConfigs.ts";

import { afterEach, describe, expect, it, vi } from "vitest";

import { fakeSheetsHttpTransport } from "../src/testSupport/fakeSheetsHttpTransport.ts";
import {
  type FakeSheetsService,
  stubSheetsService,
} from "../src/testSupport/fakeSheetsService.ts";
import {
  type ValueTypesBodyRow,
  type ValueTypesColumnName,
  valueTypesGid,
  valueTypesTableSheet,
} from "../src/testSupport/valueTypesTableSheet.ts";
import { choreSheetsConfig, confirmedSendOver } from "./choreFixture.ts";
import {
  colOf,
  type SentCase,
  type TableCase,
  topDataRowIndex,
  withoutRowIds,
  withRowIds,
} from "./restoreFixture.ts";

afterEach(() => {
  vi.unstubAllEnvs();
});

const bodyRows: ValueTypesBodyRow[] = [
  { id: "r:vty:apple", stringValue: "Apple", checkbox: true },
  { id: "r:vty:pear", stringValue: "Pear", checkbox: false },
];

// Reads the named columns, then deletes them through the framework, rightmost first so each index holds.
function deleteChoreSource(
  tableCase: TableCase,
  names: ValueTypesColumnName[],
): string {
  const colIndexes = names
    .map((name) => colOf(tableCase, name))
    .toSorted((a, b) => b - a);
  return `
export const remove = {
  description: "deletes columns of Value Types",
  action(ss) {
    ss.table("valueTypes").prepFetchColumnsFull(...${JSON.stringify(names)});
    ss.fetchAllPrepped();
    const table = ss.table("valueTypes").raw;
    for (const colIndex of ${JSON.stringify(colIndexes)}) {
      table.deleteColumn(colIndex);
    }
    ss.batchUpdateGSheets();
  },
};
`;
}

interface DeletingSend extends SentCase {
  gridBefore: unknown[][];
}

async function deletingSend(
  tableCase: TableCase,
  names: ValueTypesColumnName[],
): Promise<DeletingSend> {
  const fake = stubSheetsService({
    sheets: [valueTypesTableSheet(bodyRows, tableCase)],
  });
  const gridBefore = fake.grid.sheet(valueTypesGid).values();
  const { restore } = await confirmedSendOver({
    sheetsConfig: choreSheetsConfig({
      name: "remove",
      source: deleteChoreSource(tableCase, names),
    }),
    transport: fakeSheetsHttpTransport(fake),
    chore: "remove",
  });
  expect(fake.grid.sheet(valueTypesGid).values()).not.toEqual(gridBefore);
  return { fake, gridBefore, restore };
}

function tableColumnCount(fake: FakeSheetsService): number {
  const range = fake.grid.sheet(valueTypesGid).tables[0]?.range ?? {};
  return (range.endColumnIndex ?? 0) - (range.startColumnIndex ?? 0);
}

describe("restore of a deleting send's snapshot", () => {
  it("recreates the column where it was, head rows and column ID included, its values on their rows matched by row ID", async () => {
    const { fake, gridBefore, restore } = await deletingSend(withRowIds, [
      "stringValue",
    ]);

    const { output, exitCode } = await restore();

    expect(exitCode).toBeUndefined();
    expect(fake.grid.sheet(valueTypesGid).values()).toEqual(gridBefore);
    expect(tableColumnCount(fake)).toBe(withRowIds.columnNames.length);
    expect(output).toContain("restored:");
    expect(output).not.toContain("not restored");
  });

  it("recreates the column where it was on a Table without an ID column, its values matched by position", async () => {
    const { fake, gridBefore, restore } = await deletingSend(withoutRowIds, [
      "checkbox",
    ]);

    const { exitCode } = await restore();

    expect(exitCode).toBeUndefined();
    expect(fake.grid.sheet(valueTypesGid).values()).toEqual(gridBefore);
    expect(tableColumnCount(fake)).toBe(withoutRowIds.columnNames.length);
  });

  it("recreates the Table's first column, widening the Table again", async () => {
    const firstName = withoutRowIds.columnNames[0] ?? "stringValue";
    const { fake, gridBefore, restore } = await deletingSend(withoutRowIds, [
      firstName,
    ]);

    await restore();

    expect(fake.grid.sheet(valueTypesGid).values()).toEqual(gridBefore);
    expect(tableColumnCount(fake)).toBe(withoutRowIds.columnNames.length);
  });

  it("recreates the Table's last column, widening the Table again", async () => {
    const lastName = withRowIds.columnNames.at(-1) ?? "id";
    const { fake, gridBefore, restore } = await deletingSend(withRowIds, [
      lastName,
    ]);

    await restore();

    expect(fake.grid.sheet(valueTypesGid).values()).toEqual(gridBefore);
    expect(tableColumnCount(fake)).toBe(withRowIds.columnNames.length);
  });

  it("brings two deleted columns of one Table back in their pre-send order", async () => {
    const { fake, gridBefore, restore } = await deletingSend(withRowIds, [
      "stringValue",
      "checkbox",
    ]);

    const { output } = await restore();

    expect(fake.grid.sheet(valueTypesGid).values()).toEqual(gridBefore);
    expect(output).not.toContain("not restored");
  });
});

describe("restore of a deleting send, naming what a recreated column lacks", () => {
  it("lists each recreated column under a heading leaving its type, validation and format to version history", async () => {
    const { restore } = await deletingSend(withRowIds, [
      "stringValue",
      "checkbox",
    ]);

    const { output } = await restore();

    const lines = output.split("\n");
    const heading = lines.indexOf(
      "recreated, its column type, data validation and number format left to the spreadsheet's version history:",
    );
    expect(heading).toBeGreaterThan(-1);
    expect(lines.slice(heading + 1, heading + 3)).toEqual([
      expect.stringMatching(/^ {2}Value Types › .+$/),
      expect.stringMatching(/^ {2}Value Types › .+$/),
    ]);
  });
});

describe("restore of a deleting send when rows moved since", () => {
  it("recreates nothing without --force", async () => {
    const { fake, restore } = await deletingSend(withRowIds, ["stringValue"]);
    fake.sheets.Spreadsheets.batchUpdate(
      {
        requests: [
          {
            insertDimension: {
              range: {
                sheetId: valueTypesGid,
                dimension: "ROWS",
                startIndex: topDataRowIndex,
                endIndex: topDataRowIndex + 1,
              },
            },
          },
        ],
      },
      "spreadsheet-under-test",
    );
    const before = fake.grid.sheet(valueTypesGid).values();
    const columnCountBefore = tableColumnCount(fake);

    const { output, exitCode } = await restore();

    expect(exitCode).toBe(1);
    expect(output).toContain("RESTORE STOPPED");
    expect(fake.grid.sheet(valueTypesGid).values()).toEqual(before);
    expect(tableColumnCount(fake)).toBe(columnCountBefore);
  });
});
