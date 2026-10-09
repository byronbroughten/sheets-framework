import { describe, expect, it } from "vitest";

import { googleRawRequest } from "../00_Source/GoogleSheets/GoogleSheetsAPI";
import {
  buildGridRows,
  type FakeCell,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { SpreadsheetRaw } from "./SpreadsheetRaw";
import {
  firstTableEndRowIndex,
  gridRanges,
  lightGreen,
  tableHeaderRowIndex,
  tableId111,
  tableId222,
  topDataRowIndex,
} from "./spreadsheetRawTestSupport";

describe("SpreadsheetRaw.gatherRawOperation", () => {
  function rawValueWrite(rowIndex: number, colIndex: number, value: string) {
    return googleRawRequest({
      updateCells: {
        start: { sheetId: 111, rowIndex, columnIndex: colIndex },
        rows: [{ values: [{ userEnteredValue: { stringValue: value } }] }],
        fields: "userEnteredValue",
      },
    });
  }

  it("applies a raw request after every write the framework models, so it lands on the sheet those writes left", () => {
    const { grid } = stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          table: { endRowIndex: 11, endColumnIndex: 3 },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.gatherRawOperation(rawValueWrite(5, 2, "Raw"));
    raw.gatherRawOperation(rawValueWrite(8, 2, "Raw"));
    raw.table(tableId111).row(1).cell(2).updateValue("Processing...");
    raw.table(tableId111).row(5).cell(2).updateValue("Shifted up");
    raw.table(tableId111).row(4).delete();
    raw.batchUpdateGSheets();

    expect(
      grid.sheet(111).values({ startRowIndex: 5, endRowIndex: 10 }),
    ).toEqual([
      [null, null, "Raw"],
      [null, null, null],
      [null, null, null],
      [null, null, "Raw"],
      [null, null, null],
    ]);
  });

  it("discards a raw request alongside every other queued change, sending no batch update", () => {
    const { batchUpdateCount } = stubSheetsService();

    const raw = SpreadsheetRaw.init();
    raw.gatherRawOperation(
      googleRawRequest({ updateTable: { table: { tableId: "t" } } }),
    );
    raw.discardQueuedChanges();
    raw.batchUpdateGSheets();

    expect(batchUpdateCount()).toBe(0);
  });
});

describe("queued writes outlive a same-run re-fetch", () => {
  function stubTwoDataRows(topDataRow: FakeCell[] = ["r:lse:1", "live"]) {
    stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({
            0: ["c:lse:aaa", "c:lse:bbb"],
            [tableHeaderRowIndex]: ["ID", "Status"],
            [topDataRowIndex]: topDataRow,
            [topDataRowIndex + 1]: ["r:lse:2", "other"],
          }),
          table: { endRowIndex: topDataRowIndex + 2 },
        },
      ],
    });
  }

  function fetchedSpreadsheet() {
    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).topRow.gatherFetchFull();
    raw.table(tableId111).row(1).gatherFetchFull();
    raw.fetchAllGathered();
    return raw;
  }

  it("leaves a row queued for delete out of the working view after a re-fetch that returns it", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.table(tableId111).topRow.delete();
    raw.table(tableId111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.table(tableId111).topRow.rowInWorking()).toBe(false);
  });

  it("leaves the same row out of the working view when the re-fetch was a full row, so finalize backfilled", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.table(tableId111).topRow.delete();
    raw.table(tableId111).topRow.gatherFetchFull();
    expect(() => raw.fetchAllGathered()).not.toThrow();
    expect(raw.table(tableId111).topRow.rowInWorking()).toBe(false);
  });

  it("leaves the same row out of the working view after a full-column fetch that covers it", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.table(tableId111).topRow.delete();
    raw.table(tableId111).column(1).gatherFetchFull();
    expect(() => raw.fetchAllGathered()).not.toThrow();
    expect(raw.table(tableId111).topRow.rowInWorking()).toBe(false);
  });

  it("still supplies Table column facts from a top data row queued for delete", () => {
    stubTwoDataRows([
      "r:lse:1",
      { value: 42, isFormula: true, numberFormatType: "CURRENCY" },
    ]);

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).topRow.gatherFetchFull();
    raw.fetchAllGathered(true);
    raw.table(tableId111).topRow.delete();
    raw.table(tableId111).topRow.gatherFetchFull();
    expect(() => raw.fetchAllGathered(true)).not.toThrow();

    const { profile } = raw.table(tableId111).column(1);
    expect(profile.isFormula).toBe(true);
    expect(profile.numberFormatType).toBe("CURRENCY");
    expect(profile.topValue).toBe(42);
  });

  it("still answers whether the top data row is blank from those facts", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.table(tableId111).topRow.delete();
    raw.table(tableId111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.table(tableId111).topDataRowIsBlank()).toBe(false);
  });

  it("keeps a queued value update after a re-fetch of that cell", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.table(tableId111).topRow.cell(1).updateValue("queued");
    raw.table(tableId111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.table(tableId111).topRow.valueOrEmpty(1)).toBe("queued");
  });

  it("keeps a queued value fill after a re-fetch of a covered cell", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.table(tableId111).column(1).updateAllCells({ value: "filled" });
    raw.table(tableId111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.table(tableId111).topRow.valueOrEmpty(1)).toBe("filled");
  });

  it("lets the most recently queued value fill win when several cover one cell", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.table(tableId111).column(1).updateAllCells({ value: "first" });
    raw.table(tableId111).column(1).updateAllCells({ value: "second" });
    raw.table(tableId111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.table(tableId111).topRow.valueOrEmpty(1)).toBe("second");
  });

  it("lets the cell's own queued value win over a covering fill", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.table(tableId111).column(1).updateAllCells({ value: "filled" });
    raw.table(tableId111).topRow.cell(1).updateValue("queued");
    raw.table(tableId111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.table(tableId111).topRow.valueOrEmpty(1)).toBe("queued");
  });

  it("takes the live value when the cell has no queued value", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.table(tableId111).topRow.cell(1).updateValue("stale local");
    raw.discardQueuedChanges();
    raw.table(tableId111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.table(tableId111).topRow.valueOrEmpty(1)).toBe("live");
  });

  it("takes the live value when only a formula update is queued", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.table(tableId111).topRow.cell(1).updateFormula("=A1");
    raw.table(tableId111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.table(tableId111).topRow.valueOrEmpty(1)).toBe("live");
  });

  it("takes the live value when only a formula fill is queued", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.table(tableId111).column(1).updateAllFormulas("=A1");
    raw.table(tableId111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.table(tableId111).topRow.valueOrEmpty(1)).toBe("live");
  });

  it("integrates live values after the flush has cleared the queue", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.table(tableId111).topRow.cell(1).updateValue("queued");
    raw.batchUpdateGSheets();
    const otherRun = fetchedSpreadsheet();
    otherRun.table(tableId111).topRow.cell(1).updateValue("live");
    otherRun.batchUpdateGSheets();
    raw.table(tableId111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.table(tableId111).topRow.valueOrEmpty(1)).toBe("live");
  });

  it("applies a value queued before the row was fetched once that row is fetched", () => {
    stubTwoDataRows();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).topRow.cell(1).updateValue("queued");
    raw.table(tableId111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.table(tableId111).topRow.valueOrEmpty(1)).toBe("queued");
  });

  function stubNamedTables() {
    stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({
            0: ["c:lse:aaa", "c:lse:bbb"],
            [tableHeaderRowIndex]: ["ID", "Status"],
          }),
          table: {
            name: "records",
            endRowIndex: topDataRowIndex + 1,
            columnTypes: { 1: "TEXT" },
          },
        },
        {
          sheetId: 222,
          title: "Entries",
          rows: buildGridRows({ [tableHeaderRowIndex]: ["ID"] }),
          table: { name: "entries", endRowIndex: topDataRowIndex + 1 },
        },
      ],
    });
  }

  it("keeps a queued tab title after a re-fetch of the sheet properties", () => {
    stubNamedTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).sheet.updateTitle("Renamed");
    raw.fetchAllSheetProperties();

    expect(raw.table(tableId111).sheet.title).toBe("Renamed");
  });

  it("keeps a queued Table name on the known Table and in the sheet's Tables after a re-fetch", () => {
    stubNamedTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).updateTableName("renamedRecords");
    raw.fetchAllSheetProperties();

    expect(raw.table(tableId111).name).toBe("renamedRecords");
    expect(raw.table(tableId111).tableIds()).toEqual(["fake-table-111"]);
  });

  it("keeps a queued column type after a re-fetch of the sheet properties", () => {
    stubNamedTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).column(1).updateColumnType("DOUBLE");
    raw.fetchAllSheetProperties();

    expect(raw.table(tableId111).column(1).profile.columnType).toBe("DOUBLE");
  });

  it("lets the last of two queued tab titles win after a re-fetch", () => {
    stubNamedTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).sheet.updateTitle("First");
    raw.table(tableId111).sheet.updateTitle("Second");
    raw.fetchAllSheetProperties();

    expect(raw.table(tableId111).sheet.title).toBe("Second");
  });

  it("lets the last of two queued Table names win after a re-fetch", () => {
    stubNamedTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).updateTableName("firstRecords");
    raw.table(tableId111).updateTableName("secondRecords");
    raw.fetchAllSheetProperties();

    expect(raw.table(tableId111).name).toBe("secondRecords");
  });

  it("leaves another sheet's queued title alone on a re-fetch of one sheet", () => {
    stubNamedTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId222).sheet.updateTitle("Renamed");
    raw.fetchSheetUsedGrid(111);

    expect(raw.table(tableId222).sheet.title).toBe("Renamed");
  });

  it("applies a sheet's queued title to that sheet only", () => {
    stubNamedTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).sheet.updateTitle("Renamed");
    raw.fetchAllSheetProperties();

    expect(raw.table(tableId222).sheet.title).toBe("Entries");
  });

  it("integrates the live title and Table name after the flush has cleared the queue", () => {
    stubNamedTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).sheet.updateTitle("Renamed");
    raw.table(tableId111).updateTableName("renamedRecords");
    raw.batchUpdateGSheets();
    const otherRun = SpreadsheetRaw.init();
    otherRun.fetchAllSheetProperties();
    otherRun.table(tableId111).sheet.updateTitle("Records");
    otherRun.table(tableId111).updateTableName("records");
    otherRun.batchUpdateGSheets();
    raw.fetchAllSheetProperties();

    expect(raw.table(tableId111).sheet.title).toBe("Records");
    expect(raw.table(tableId111).name).toBe("records");
  });
});

describe("the last queued write wins between fills and cell writes", () => {
  function stubFilledSheet() {
    return stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({
            0: ["c:lse:aaa", "c:lse:bbb"],
            4: ["r:lse:1", "old"],
            5: ["r:lse:2", "old"],
          }),
          table: { endRowIndex: 6 },
        },
      ],
    });
  }
  function fetchedColumn() {
    const raw = SpreadsheetRaw.init();
    raw.table(tableId111, 111).column(1).gatherFetchFull();
    raw.fetchAllGathered();
    return raw;
  }

  it("shows a fill queued after a cell write", () => {
    const { grid } = stubFilledSheet();

    const raw = fetchedColumn();
    raw.table(tableId111).row(1).cell(1).updateValue("cell");
    raw.table(tableId111).column(1).updateAllCells({ value: "filled" });
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).values(gridRanges.columnOneData)).toEqual([
      ["filled"],
      ["filled"],
    ]);
  });

  it("keeps the cell's colour under a later value-only fill", () => {
    const { grid } = stubFilledSheet();

    const raw = fetchedColumn();
    raw
      .table(tableId111)
      .row(1)
      .cell(1)
      .updateValue("cell")
      .updateBackgroundColor(lightGreen);
    raw.table(tableId111).column(1).updateAllCells({ value: "filled" });
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).rows(gridRanges.columnOneData)).toEqual([
      ["filled"],
      [{ value: "filled", backgroundColor: lightGreen }],
    ]);
  });

  it("keeps the cell's value under a later colour-only fill", () => {
    const { grid } = stubFilledSheet();

    const raw = fetchedColumn();
    raw.table(tableId111).row(1).cell(1).updateValue("cell");
    raw
      .table(tableId111)
      .column(1)
      .updateAllCells({ backgroundColor: lightGreen });
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).rows(gridRanges.columnOneData)).toEqual([
      [{ value: "old", backgroundColor: lightGreen }],
      [{ value: "cell", backgroundColor: lightGreen }],
    ]);
  });

  it("drops a queued cell formula under a later value fill", () => {
    const { grid } = stubFilledSheet();

    const raw = fetchedColumn();
    raw.table(tableId111).row(1).cell(1).updateFormula("=1+1");
    raw.table(tableId111).column(1).updateAllCells({ value: "filled" });
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).rows(gridRanges.columnOneData)).toEqual([
      ["filled"],
      ["filled"],
    ]);
  });

  it("reads before the flush what the grid shows after it, across a re-fetch", () => {
    const { grid } = stubFilledSheet();

    const raw = fetchedColumn();
    raw.table(tableId111).row(1).cell(1).updateValue("cell");
    raw.table(tableId111).column(1).updateAllCells({ value: "filled" });
    raw.table(tableId111).row(1).gatherFetchFull();
    raw.fetchAllGathered();
    const readBeforeFlush = raw.table(tableId111).row(1).valueOrEmpty(1);
    raw.batchUpdateGSheets();

    expect([readBeforeFlush, grid.sheet(111).cell(5, 1)]).toEqual([
      "filled",
      "filled",
    ]);
  });

  it("leaves a cell write on a row appended after the fill alone", () => {
    const { grid } = stubFilledSheet();

    const raw = fetchedColumn();
    raw.table(tableId111).column(1).updateAllCells({ value: "filled" });
    raw.table(tableId111).appendDataRow().cell(1).updateValue("appended");
    raw.batchUpdateGSheets();

    expect(
      grid.sheet(111).values({ ...gridRanges.columnOneData, endRowIndex: 7 }),
    ).toEqual([["filled"], ["filled"], ["appended"]]);
  });
});

describe("SpreadsheetRaw.discardQueuedChanges", () => {
  it("sends no batch update for changes queued before the discard", () => {
    const { batchUpdateCount } = stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).row(1).delete();
    raw.discardQueuedChanges();
    raw.batchUpdateGSheets();

    expect(batchUpdateCount()).toBe(0);
    expect(raw.table(tableId111).rowIndexesAreStale).toBe(false);
    expect(raw.table(tableId111).dataRowCount).toBe(7);
  });

  it("empties the spreadsheet and per-sheet write queues, so a later flush sends no batch update", () => {
    const { batchUpdateCount } = stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).row(1).delete();
    raw.table(tableId111).requestSortGSheet({
      colIdxToSortBy: 0,
      sortOrder: "ASCENDING",
    });
    raw.gatherRawOperation(
      googleRawRequest({ updateTable: { table: { tableId: "t" } } }),
    );
    raw.discardQueuedChanges();
    raw.batchUpdateGSheets();

    expect(batchUpdateCount()).toBe(0);
  });

  it("still applies changes queued after the discard, so a failure handler can report status", () => {
    const { grid } = stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({ 4: ["r4"], 5: ["r5"] }),
          table: { endRowIndex: 11 },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).row(1).delete();
    raw.discardQueuedChanges();
    raw.table(tableId111).appendDataRow();
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).values({ startRowIndex: 4 })).toEqual([
      ["r4"],
      ["r5"],
    ]);
    expect(firstTableEndRowIndex(grid, 111)).toBe(12);
  });
});
