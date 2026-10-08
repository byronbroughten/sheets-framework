import { describe, expect, it } from "vitest";

import {
  buildGridRows,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { SpreadsheetRaw } from "./SpreadsheetRaw";
import {
  lightGreen,
  tableId111,
  threeByThreeAround,
} from "./spreadsheetRawTestSupport";

describe("CellRaw.updateValue", () => {
  it("writes to a row that was never fetched, since a write needs no fetched state", () => {
    const { grid } = stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).row(1).cell(2).updateValue("Processing...");
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).values()).toEqual([
      [null, null, null],
      [null, null, null],
      [null, null, null],
      [null, null, null],
      [null, null, null],
      [null, null, "Processing..."],
    ]);
  });

  it("leaves an unfetched row unreadable, so a forgotten fetch still fails loudly on read", () => {
    stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    const cell = raw.table(tableId111).row(1).cell(2);
    cell.updateValue("Processing...");

    expect(() => cell.valueOrEmpty()).toThrowError(/No value is set/);
  });

  it("throws for a data row past the table's last row rather than writing off the grid", () => {
    stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();

    expect(() =>
      raw.table(tableId111).row(7).cell(2).updateValue("x"),
    ).toThrowError(/past the last row/);
  });

  it("reflects the write in row state when the row was fetched, so a later read sees it", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({
            0: ["c:lse:aaa", "c:lse:bbb"],
            4: ["r:lse:1", "old"],
          }),
          table: { endRowIndex: 5 },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).row(0).gatherFetchFull();
    raw.fetchAllGathered();
    const cell = raw.table(tableId111).row(0).cell(1);
    cell.updateValue("new");

    expect(cell.valueOrEmpty()).toBe("new");
  });
});

describe("CellRaw.updateBackgroundColor", () => {
  it("colours the cell and leaves its value alone", () => {
    const { grid } = stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({ 5: [null, null, "existing"] }),
          table: { endRowIndex: 11 },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).row(1).cell(2).updateBackgroundColor(lightGreen);
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).cell(5, 2)).toEqual({
      value: "existing",
      backgroundColor: lightGreen,
    });
  });

  it("writes the value queued before the colour on the same cell", () => {
    const { grid } = stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).row(1).cell(2).updateValue("2026-09-05 10:00:00");
    raw.table(tableId111).row(1).cell(2).updateBackgroundColor(lightGreen);
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).cell(5, 2)).toEqual({
      value: "2026-09-05 10:00:00",
      backgroundColor: lightGreen,
    });
  });

  it("leaves the cell unreadable, since the read path never fetches colour", () => {
    stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    const cell = raw.table(tableId111).row(1).cell(2);
    cell.updateBackgroundColor(lightGreen);

    expect(cell.inWorking).toBe(false);
    expect(() => cell.valueOrEmpty()).toThrowError(/No value is set/);
  });
});

describe("CellRaw.addCheckboxValidation", () => {
  it("makes that one cell a checkbox", () => {
    const { grid } = stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).row(1).cell(2).addCheckboxValidation();
    raw.batchUpdateGSheets();

    const checkbox = { value: null, dataValidationConditionType: "BOOLEAN" };
    expect(grid.sheet(111).rows(threeByThreeAround(5, 2))).toEqual([
      [null, null, null],
      [null, checkbox, null],
      [null, null, null],
    ]);
  });
});

// A mis-wired accessor still type-checks; the instance checks catch it.
