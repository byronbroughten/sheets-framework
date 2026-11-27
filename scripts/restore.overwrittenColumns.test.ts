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
import {
  choreSheetsConfig,
  confirmedSendOver,
  srcModule,
} from "./choreFixture.ts";
import {
  colOf,
  type SentCase,
  type TableCase,
  topDataRowIndex,
  withoutRowIds,
  withRowIds,
} from "./restoreFixture.ts";

type GoogleRequest = GoogleAppsScript.Sheets.Schema.Request;

afterEach(() => {
  vi.unstubAllEnvs();
});

const bodyRows: ValueTypesBodyRow[] = [
  { id: "r:vty:apple", stringValue: "Apple", checkbox: true },
  { id: "r:vty:pear", stringValue: "Pear", checkbox: true },
];

function writeCell(
  rowIndex: number,
  colIndex: number,
  userEnteredValue: GoogleAppsScript.Sheets.Schema.ExtendedValue,
): GoogleRequest {
  return {
    updateCells: {
      range: {
        sheetId: valueTypesGid,
        startRowIndex: rowIndex,
        endRowIndex: rowIndex + 1,
        startColumnIndex: colIndex,
        endColumnIndex: colIndex + 1,
      },
      rows: [{ values: [{ userEnteredValue }] }],
      fields: "userEnteredValue",
    },
  };
}

// Reads Text and Checkbox, then overwrites both on every row, which the gate holds.
function overwriteChoreSource(
  tableCase: TableCase,
  rows: ValueTypesBodyRow[],
): string {
  const requests = rows.flatMap((_, bodyIndex) => [
    writeCell(topDataRowIndex + bodyIndex, colOf(tableCase, "stringValue"), {
      stringValue: "Overwritten",
    }),
    writeCell(topDataRowIndex + bodyIndex, colOf(tableCase, "checkbox"), {
      boolValue: false,
    }),
  ]);
  return `
import { googleRawRequest } from ${srcModule("00_Source/GoogleSheets/GoogleSheetsAPI.ts")};
import { SpreadsheetRaw } from ${srcModule("02_SpreadsheetRaw/SpreadsheetRaw.ts")};

export const overwrite = {
  description: "overwrites Text and Checkbox on every row",
  action(ss) {
    ss.table("valueTypes").prepFetchColumnsFull("stringValue", "checkbox");
    ss.fetchAllPrepped();
    const raw = SpreadsheetRaw.init();
    for (const request of ${JSON.stringify(requests)}) {
      raw.gatherRawOperation(googleRawRequest(request));
    }
    raw.batchUpdateGSheets();
  },
};
`;
}

// Spare blank rows below the Table, as a live tab keeps; the fake otherwise sizes its grid to the Table.
async function overwritingSend(
  tableCase: TableCase,
  rows: ValueTypesBodyRow[] = bodyRows,
  spareRowCount = 0,
): Promise<SentCase> {
  const sheet = valueTypesTableSheet(rows, tableCase);
  const spareRows = Array.from({ length: spareRowCount }, () => []);
  const fake = stubSheetsService({
    sheets: [{ ...sheet, rows: [...(sheet.rows ?? []), ...spareRows] }],
  });
  const { restore } = await confirmedSendOver({
    sheetsConfig: choreSheetsConfig({
      name: "overwrite",
      source: overwriteChoreSource(tableCase, rows),
    }),
    transport: fakeSheetsHttpTransport(fake),
    chore: "overwrite",
  });
  return { fake, restore };
}

function columnValues(
  fake: FakeSheetsService,
  tableCase: TableCase,
  name: ValueTypesColumnName,
): unknown[] {
  return fake.grid
    .sheet(valueTypesGid)
    .values()
    .slice(topDataRowIndex)
    .map((row) => row[colOf(tableCase, name)]);
}

function edit(fake: FakeSheetsService, requests: GoogleRequest[]): void {
  fake.sheets.Spreadsheets.batchUpdate({ requests }, "spreadsheet-under-test");
}

function insertRowAbove(rowIndex: number): GoogleRequest {
  return {
    insertDimension: {
      range: {
        sheetId: valueTypesGid,
        dimension: "ROWS",
        startIndex: rowIndex,
        endIndex: rowIndex + 1,
      },
    },
  };
}

function deleteRow(rowIndex: number): GoogleRequest {
  return {
    deleteDimension: {
      range: {
        sheetId: valueTypesGid,
        dimension: "ROWS",
        startIndex: rowIndex,
        endIndex: rowIndex + 1,
      },
    },
  };
}

describe("restore of an overwriting send's snapshot", () => {
  it("puts each saved column back on its rows, matched by row ID", async () => {
    const { fake, restore } = await overwritingSend(withRowIds);
    expect(columnValues(fake, withRowIds, "stringValue")).toEqual([
      "Overwritten",
      "Overwritten",
    ]);

    const { output, exitCode } = await restore();

    expect(exitCode).toBeUndefined();
    expect(columnValues(fake, withRowIds, "stringValue")).toEqual([
      "Apple",
      "Pear",
    ]);
    expect(columnValues(fake, withRowIds, "checkbox")).toEqual([true, true]);
    expect(output).toContain("restored:");
  });

  it("puts each saved column back on its rows, matched by position, on a Table without an ID column", async () => {
    const { fake, restore } = await overwritingSend(withoutRowIds);

    const { exitCode } = await restore();

    expect(exitCode).toBeUndefined();
    expect(columnValues(fake, withoutRowIds, "stringValue")).toEqual([
      "Apple",
      "Pear",
    ]);
    expect(columnValues(fake, withoutRowIds, "checkbox")).toEqual([true, true]);
  });

  it("lists no recreated column when no column was deleted", async () => {
    const { restore } = await overwritingSend(withRowIds);

    const { output } = await restore();

    expect(output).not.toContain(
      "recreated, its column type, data validation and number format left to the spreadsheet's version history:",
    );
  });

  it("puts a saved row with a blank row ID back by position when no row moved", async () => {
    const { fake, restore } = await overwritingSend(withRowIds, [
      { id: "r:vty:apple", stringValue: "Apple", checkbox: true },
      { stringValue: "Pear", checkbox: true },
    ]);

    const { output, exitCode } = await restore();

    expect(exitCode).toBeUndefined();
    expect(columnValues(fake, withRowIds, "stringValue")).toEqual([
      "Apple",
      "Pear",
    ]);
    expect(output).not.toContain("left to the spreadsheet's version history");
  });
});

describe("restore when rows moved since the send", () => {
  it("writes nothing, exits nonzero and lists the changed tab", async () => {
    const { fake, restore } = await overwritingSend(withRowIds);
    edit(fake, [insertRowAbove(topDataRowIndex)]);
    const before = fake.grid.sheet(valueTypesGid).values();

    const { output, exitCode } = await restore();

    expect(exitCode).toBe(1);
    expect(output).toContain("RESTORE STOPPED");
    expect(output).toContain("  Value Types");
    expect(fake.grid.sheet(valueTypesGid).values()).toEqual(before);
  });

  it("restores anyway with --force, each value onto its row ID's new row", async () => {
    const { fake, restore } = await overwritingSend(withRowIds);
    edit(fake, [insertRowAbove(topDataRowIndex)]);

    const { output, exitCode } = await restore(["--force"]);

    expect(exitCode).toBeUndefined();
    expect(output).toContain("restored anyway");
    expect(columnValues(fake, withRowIds, "id")).toEqual([
      null,
      "r:vty:apple",
      "r:vty:pear",
    ]);
    expect(columnValues(fake, withRowIds, "stringValue")).toEqual([
      null,
      "Apple",
      "Pear",
    ]);
  });

  it("lists a saved row whose row ID is gone as left to version history", async () => {
    const { fake, restore } = await overwritingSend(withRowIds);
    edit(fake, [deleteRow(topDataRowIndex)]);

    const { output } = await restore(["--force"]);

    expect(columnValues(fake, withRowIds, "stringValue")).toEqual(["Pear"]);
    expect(output).toContain(
      "left to the spreadsheet's version history, with no row to land on:",
    );
    expect(output).toContain("row ID r:vty:apple");
  });

  it("lists a saved row past the Table's last row as left to version history, on a Table without an ID column", async () => {
    const { fake, restore } = await overwritingSend(withoutRowIds);
    edit(fake, [deleteRow(topDataRowIndex + 1)]);

    const { output } = await restore(["--force"]);

    expect(columnValues(fake, withoutRowIds, "stringValue")).toEqual(["Apple"]);
    expect(output).toContain(`row ${topDataRowIndex + 2}`);
  });

  it("lists a saved row past the Table's last row as left to version history, never writing it onto a spare blank row below the Table", async () => {
    const { fake, restore } = await overwritingSend(withoutRowIds, bodyRows, 5);
    edit(fake, [deleteRow(topDataRowIndex + 1)]);
    const rowCountBefore = fake.grid.sheet(valueTypesGid).rowCount;

    const { output } = await restore(["--force"]);

    const stringValues = columnValues(fake, withoutRowIds, "stringValue");
    expect(stringValues[0]).toBe("Apple");
    expect(stringValues).not.toContain("Pear");
    expect(columnValues(fake, withoutRowIds, "checkbox")).toEqual([
      true,
      ...stringValues.slice(1).map(() => null),
    ]);
    expect(fake.grid.sheet(valueTypesGid).rowCount).toBe(rowCountBefore);
    expect(output).toContain(
      "left to the spreadsheet's version history, with no row to land on:",
    );
    expect(output).toContain(`row ${topDataRowIndex + 2}`);
  });
});

describe("restore when a saved column's tab is gone", () => {
  it("lists each of its columns as not restored, with --force", async () => {
    const { fake, restore } = await overwritingSend(withRowIds);
    edit(fake, [{ deleteSheet: { sheetId: valueTypesGid } }]);

    const { output, exitCode } = await restore(["--force"]);

    expect(exitCode).toBeUndefined();
    expect(output).toContain("not restored, its tab or Table is gone:");
    expect(output).toContain("  Value Types › Checkbox (c:vty:checkbox)");
  });
});
