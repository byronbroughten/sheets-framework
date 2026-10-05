import { describe, expect, it } from "vitest";

import { installedRawSource } from "../00_Source/RawSource/RawSource";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  buildGridRows,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { SpreadsheetRaw } from "./SpreadsheetRaw";
import {
  colIdRowIndex,
  extraTablesSheet,
  firstTableEndRowIndex,
  itemGid,
  lightGreen,
  logGid,
  misplacedTableSheet,
  placedTableSheet,
  recordedGridRanges,
  scratchGid,
  startTableColIndex,
  tableEndRowIndex,
  tableHeaderRowIndex,
  thrownMessage,
  topDataRowIndex,
} from "./spreadsheetRawTestSupport";

describe("SpreadsheetRaw.fetchAllSheetProperties", () => {
  it("integrates sheet properties from Sheets.Spreadsheets.get into raw state", () => {
    stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records" }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();

    expect(raw.activeSheetGids).toEqual([111]);
    expect(raw.sheet(111).title).toBe("Records");
  });

  it("throws when a known sheet has more than one Table on the unfiltered census", () => {
    stubSheetsService({
      sheets: [extraTablesSheet({ sheetId: itemGid, title: "Item" })],
    });

    const raw = SpreadsheetRaw.init();
    expect(() => raw.fetchAllSheetProperties()).toThrowError(
      /1 sheet\(s\) have more than one Table — delete the extras so each sheet has exactly one: "Item" \(gid \d+\)/,
    );
  });
});

describe("SpreadsheetRaw.timeZone", () => {
  it("fetches the zone alone, once, and logs it when nothing has been fetched", () => {
    const logger = stubLogger();
    const { getCalls } = stubSheetsService({ timeZone: "Europe/London" });

    const raw = SpreadsheetRaw.init();

    expect(raw.timeZone).toBe("Europe/London");
    expect(raw.timeZone).toBe("Europe/London");
    expect(getCalls).toEqual([{ fields: "properties(timeZone)" }]);
    expect(logger.log).toHaveBeenCalledTimes(1);
  });

  it("reads the zone that rode a grid fetch, silently and with no get of its own", () => {
    const logger = stubLogger();
    const { getCalls } = stubSheetsService({
      timeZone: "Australia/Sydney",
      sheets: [placedTableSheet({ sheetId: 111, title: "Records" })],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchSheetUsedGrid(111);

    expect(raw.timeZone).toBe("Australia/Sydney");
    expect(getCalls).toEqual([]);
    expect(logger.log).not.toHaveBeenCalled();
  });

  it("throws when the spreadsheet's time zone is absent", () => {
    stubLogger();
    stubSheetsService({ timeZone: null });

    const raw = SpreadsheetRaw.init();

    expect(() => raw.timeZone).toThrowError(/properties\.timeZone/);
  });
});

describe("SpreadsheetRaw.fetchAllGathered", () => {
  it("throws one aggregate error naming every sheet queued for a full fetch that has no Table", () => {
    stubSheetsService({
      sheets: [
        { sheetId: 111, title: "Task Generic" },
        { sheetId: 222, title: "Task Material" },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(111).gatherFetchProperties();
    raw.sheetMeta(111).gatherFetchColumnIdsInit();
    raw.sheet(222).gatherFetchProperties();
    raw.sheetMeta(222).gatherFetchColumnIdsInit();

    expect(() => raw.fetchAllGathered()).toThrowError(
      /"Task Generic" \(gid 111\).*"Task Material" \(gid 222\)/,
    );
  });

  it("does not throw for a sheet with a Table", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({ 0: ["ID"] }),
          table: { endRowIndex: 5 },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(111).gatherFetchProperties();

    expect(() => raw.fetchAllGathered()).not.toThrow();
  });

  it("does not throw for a config-known sheet whose Table starts where the layout requires", () => {
    stubSheetsService({
      sheets: [placedTableSheet({ sheetId: itemGid, title: "Item" })],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(itemGid).gatherFetchProperties();

    expect(() => raw.fetchAllGathered()).not.toThrow();
  });

  it("names both the found and the required position for a Table one row too high", () => {
    stubSheetsService({
      sheets: [
        misplacedTableSheet({
          sheetId: itemGid,
          title: "Item",
          startRowIndex: tableHeaderRowIndex - 1,
        }),
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(itemGid).gatherFetchProperties();

    expect(() => raw.fetchAllGathered()).toThrowError(
      /"Item".*starts at row 3, column A.*must start at row 4, column A/,
    );
  });

  it("names both positions for a Table one column to the right of the layout", () => {
    stubSheetsService({
      sheets: [
        misplacedTableSheet({
          sheetId: itemGid,
          title: "Item",
          startColumnIndex: startTableColIndex + 1,
        }),
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(itemGid).gatherFetchProperties();

    expect(() => raw.fetchAllGathered()).toThrowError(
      /"Item".*starts at row 4, column B.*must start at row 4, column A/,
    );
  });

  it("leaves a sheet the config does not know alone, however its Table is placed", () => {
    stubSheetsService({
      sheets: [
        misplacedTableSheet({
          sheetId: scratchGid,
          title: "Byron's Scratch Sheet",
          startRowIndex: tableHeaderRowIndex - 1,
        }),
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(scratchGid).gatherFetchProperties();

    expect(() => raw.fetchAllGathered()).not.toThrow();
  });

  it("names every misplaced sheet in one error, including one nothing was queued for", () => {
    stubSheetsService({
      sheets: [
        misplacedTableSheet({
          sheetId: itemGid,
          title: "Item",
          startRowIndex: tableHeaderRowIndex - 1,
        }),
        misplacedTableSheet({
          sheetId: logGid,
          title: "Log",
          startColumnIndex: startTableColIndex + 1,
        }),
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(itemGid).gatherFetchProperties();

    expect(() => raw.fetchAllGathered()).toThrowError(/"Item".*"Log"/);
  });

  it("reports a Table the filtered fetch could not see as misplaced rather than absent", () => {
    stubSheetsService({
      sheets: [
        {
          ...misplacedTableSheet({
            sheetId: itemGid,
            title: "Item",
            startRowIndex: tableHeaderRowIndex + 2,
          }),
          isTableHiddenFromFilteredFetch: true,
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(itemGid).gatherFetchProperties();
    raw.sheetMeta(itemGid).gatherFetchColumnIdsInit();

    const message = thrownMessage(() => raw.fetchAllGathered());
    expect(message).toMatch(
      /"Item".*starts at row 6, column A.*must start at row 4, column A/,
    );
    expect(message).not.toMatch(/Insert > Table/);
  });

  it("throws naming a known sheet whose gathered payload has more than one Table, and does not keep the first as active", () => {
    stubSheetsService({
      sheets: [extraTablesSheet({ sheetId: itemGid, title: "Item" })],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(itemGid).gatherFetchProperties();

    const message = thrownMessage(() => raw.fetchAllGathered());
    expect(message).toMatch(
      /1 sheet\(s\) have more than one Table — delete the extras so each sheet has exactly one: "Item" \(gid \d+\)/,
    );
    expect(raw.sheet(itemGid).hasFetchedProperties).toBe(false);
  });

  it("names every known sheet with extra Tables in one error", () => {
    stubSheetsService({
      sheets: [
        extraTablesSheet({ sheetId: itemGid, title: "Item" }),
        extraTablesSheet({ sheetId: logGid, title: "Log" }),
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(itemGid).gatherFetchProperties();
    raw.sheet(logGid).gatherFetchProperties();

    expect(() => raw.fetchAllGathered()).toThrowError(/"Item".*"Log"/);
  });

  it("leaves a sheet the config does not know alone, even with two Tables", () => {
    stubSheetsService({
      sheets: [
        extraTablesSheet({
          sheetId: scratchGid,
          title: "Byron's Scratch Sheet",
        }),
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(scratchGid).gatherFetchProperties();

    expect(() => raw.fetchAllGathered()).not.toThrow();
  });

  it("names extra Tables and a missing Table in one error", () => {
    stubSheetsService({
      sheets: [
        extraTablesSheet({ sheetId: itemGid, title: "Item" }),
        { sheetId: logGid, title: "Log" },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(itemGid).gatherFetchProperties();
    raw.sheet(logGid).gatherFetchProperties();
    raw.sheetMeta(logGid).gatherFetchColumnIdsInit();

    const message = thrownMessage(() => raw.fetchAllGathered());
    expect(message).toMatch(/more than one Table.*"Item"/);
    expect(message).toMatch(/Insert > Table.*"Log"/);
  });

  it("names extra Tables and a misplaced Table in one error", () => {
    stubSheetsService({
      sheets: [
        extraTablesSheet({ sheetId: itemGid, title: "Item" }),
        misplacedTableSheet({
          sheetId: logGid,
          title: "Log",
          startRowIndex: tableHeaderRowIndex - 1,
        }),
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(itemGid).gatherFetchProperties();
    raw.sheet(logGid).gatherFetchProperties();

    const message = thrownMessage(() => raw.fetchAllGathered());
    expect(message).toMatch(/more than one Table.*"Item"/);
    expect(message).toMatch(/does not start where the layout requires.*"Log"/);
  });

  it("reports extra Tables the filtered fetch could not see as extras rather than absent", () => {
    stubSheetsService({
      sheets: [
        {
          ...extraTablesSheet({ sheetId: itemGid, title: "Item" }),
          isTableHiddenFromFilteredFetch: true,
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(itemGid).gatherFetchProperties();
    raw.sheetMeta(itemGid).gatherFetchColumnIdsInit();

    const message = thrownMessage(() => raw.fetchAllGathered());
    expect(message).toMatch(/more than one Table.*"Item"/);
    expect(message).not.toMatch(/Insert > Table/);
  });

  it("sends no request when no ranges were gathered, since empty dataFilters would fetch the whole spreadsheet", () => {
    const { getByDataFilterCalls } = stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({ 0: ["ID"] }),
          table: { endRowIndex: 5 },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllGathered();

    expect(getByDataFilterCalls).toEqual([]);
    expect(raw.activeSheetGids).toEqual([]);
  });

  it("aims the properties probe at one header cell on the layout start column", () => {
    const { getByDataFilterCalls } = stubSheetsService({
      sheets: [placedTableSheet({ sheetId: itemGid, title: "Item" })],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(itemGid).gatherFetchProperties();
    raw.fetchAllGathered();

    expect(recordedGridRanges(getByDataFilterCalls)).toEqual([
      {
        sheetId: itemGid,
        startRowIndex: tableHeaderRowIndex,
        endRowIndex: tableHeaderRowIndex + 1,
        startColumnIndex: startTableColIndex,
        endColumnIndex: startTableColIndex + 1,
      },
    ]);
  });

  it("aims the column-id filter at the layout start column", () => {
    const { getByDataFilterCalls } = stubSheetsService({
      sheets: [placedTableSheet({ sheetId: itemGid, title: "Item" })],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheetMeta(itemGid).gatherFetchColumnIdsInit();
    raw.fetchAllGathered();

    expect(recordedGridRanges(getByDataFilterCalls)).toEqual([
      {
        sheetId: itemGid,
        startRowIndex: colIdRowIndex,
        endRowIndex: colIdRowIndex + 1,
        startColumnIndex: startTableColIndex,
      },
    ]);
  });

  it("refuses a full-row fetch before the sheet has a Table in state", () => {
    stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records" }],
    });

    const raw = SpreadsheetRaw.init();

    expect(() => raw.sheet(111).topRow.gatherFetchFull()).toThrowError(
      /Table is unknown for sheetGid 111/,
    );
  });

  it("aims a full-row fetch at the live Table's own columns after properties, not the layout constant", () => {
    const liveStart = startTableColIndex + 1;
    const { getByDataFilterCalls } = stubSheetsService({
      sheets: [
        {
          sheetId: scratchGid,
          title: "Byron's Scratch Sheet",
          rows: buildGridRows({
            [tableHeaderRowIndex]: ["", "ID", "Name"],
          }),
          table: {
            startColumnIndex: liveStart,
            endColumnIndex: liveStart + 2,
            endRowIndex: tableEndRowIndex,
          },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(scratchGid).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(recordedGridRanges(getByDataFilterCalls)).toEqual([
      {
        sheetId: scratchGid,
        startRowIndex: topDataRowIndex,
        endRowIndex: topDataRowIndex + 1,
        startColumnIndex: liveStart,
        endColumnIndex: liveStart + 2,
      },
    ]);
  });

  it("aims a full-column fetch from the column ID row to the Table's last row", () => {
    const { getByDataFilterCalls } = stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({
            [tableHeaderRowIndex]: ["ID", "Name"],
            [topDataRowIndex]: ["r1", "a"],
            [topDataRowIndex + 1]: ["r2", "b"],
          }),
          table: { endRowIndex: topDataRowIndex + 2 },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).column(1).gatherFetchFull();
    raw.fetchAllGathered();

    expect(recordedGridRanges(getByDataFilterCalls)).toEqual([
      {
        sheetId: 111,
        startRowIndex: colIdRowIndex,
        endRowIndex: topDataRowIndex + 2,
        startColumnIndex: startTableColIndex + 1,
        endColumnIndex: startTableColIndex + 2,
      },
    ]);
    expect(raw.sheet(111).column(1).valueArrOrEmpty).toEqual(["a", "b"]);
  });

  it("fetches a full row, a full column and a cell in one round trip", () => {
    const { getByDataFilterCalls } = stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({
            [tableHeaderRowIndex]: ["ID", "Name"],
            [topDataRowIndex]: ["r1", "a"],
            [topDataRowIndex + 1]: ["r2", "b"],
          }),
          table: { endRowIndex: topDataRowIndex + 2 },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    const sheet = raw.sheet(111);
    sheet.topRow.gatherFetchFull();
    sheet.column(1).gatherFetchFull();
    sheet.row(1).cell(0).gatherFetchRange();
    raw.fetchAllGathered();

    expect(getByDataFilterCalls).toHaveLength(1);
    expect(sheet.row(0).valueOrEmpty(0)).toBe("r1");
    expect(sheet.row(1).valueOrEmpty(0)).toBe("r2");
    expect(sheet.row(1).valueOrEmpty(1)).toBe("b");
  });
});

describe("SpreadsheetRaw over a Table placed lower on its sheet", () => {
  // Header on sheet row 9 (gutter 10), first column C; a loose cell sits outside the Table.
  const lowerHeaderRowIndex = 9;
  const lowerStartColIndex = 2;

  function stubLowerTable() {
    return stubSheetsService({
      sheets: [
        {
          sheetId: scratchGid,
          title: "Lower",
          rows: buildGridRows({
            [lowerHeaderRowIndex + 1]: ["loose", null, "r:low:1", "first"],
            [lowerHeaderRowIndex + 2]: [null, null, "r:low:2", "second"],
          }),
          table: {
            startRowIndex: lowerHeaderRowIndex,
            startColumnIndex: lowerStartColIndex,
            endRowIndex: lowerHeaderRowIndex + 3,
            endColumnIndex: lowerStartColIndex + 2,
            headRows: {
              0: ["ID", "Name"],
              3: ["c:low:aaa", "c:low:bbb"],
            },
          },
        },
      ],
    });
  }

  function fetchedLowerRaw(): SpreadsheetRaw {
    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheetMeta(scratchGid).gatherFetchColumnIdsInit();
    raw.sheet(scratchGid).row(0).gatherFetchFull();
    raw.sheet(scratchGid).row(1).gatherFetchFull();
    raw.fetchAllGathered();
    return raw;
  }

  it("reads body row 0 from the row just below the header, and column 0 from the Table's first column", () => {
    stubLowerTable();
    const raw = fetchedLowerRaw();
    const sheet = raw.sheet(scratchGid);

    expect(sheet.rowIndexesFull).toEqual([0, 1]);
    expect(sheet.row(0).valueOrEmpty(0)).toBe("r:low:1");
    expect(sheet.row(1).valueOrEmpty(1)).toBe("second");
    expect(raw.sheetMeta(scratchGid).activeColumnIds).toEqual([
      "c:low:aaa",
      "c:low:bbb",
    ]);
    expect(raw.sheetMeta(scratchGid).tableHeaderRow.valueOrEmpty(1)).toBe(
      "Name",
    );
  });

  it("updates and appends at the Table's grid rows, leaving the loose cell beside it alone", () => {
    const { grid } = stubLowerTable();
    const raw = fetchedLowerRaw();
    const sheet = raw.sheet(scratchGid);

    sheet.row(1).cell(1).updateValue("changed");
    sheet.appendDataRow().cell(0).updateValue("r:low:3");
    sheet.column(1).updateActiveCells({ backgroundColor: lightGreen });
    raw.batchUpdateGSheets();

    expect(
      grid.sheet(scratchGid).values({
        startRowIndex: lowerHeaderRowIndex,
        endRowIndex: lowerHeaderRowIndex + 4,
      }),
    ).toEqual([
      [null, null, "ID", "Name"],
      ["loose", null, "r:low:1", "first"],
      [null, null, "r:low:2", "changed"],
      [null, null, "r:low:3", null],
    ]);
    expect(firstTableEndRowIndex(grid, scratchGid)).toBe(
      lowerHeaderRowIndex + 4,
    );
    expect(grid.sheet(scratchGid).cell(lowerHeaderRowIndex + 1, 3)).toEqual({
      value: "first",
      backgroundColor: lightGreen,
    });
  });

  it("names a row in a message by the gutter number the operator sees", () => {
    stubLowerTable();
    const raw = fetchedLowerRaw();

    expect(() => raw.sheet(scratchGid).row(5).valueOrEmpty(0)).toThrowError(
      "No value is set in row 16 for column index 0.",
    );
  });
});

describe("SpreadsheetRaw fetch integration routes each cell to its Table", () => {
  const leftStart = startTableColIndex;
  const rightStart = startTableColIndex + 3;
  function stubSideBySideTables(tableIds = ["left", "right"]) {
    return stubSheetsService({
      sheets: [
        {
          sheetId: scratchGid,
          title: "Byron's Scratch Sheet",
          rows: buildGridRows({
            [colIdRowIndex - 1]: ["loose-above"],
            [tableHeaderRowIndex]: ["ID", "Name", "", "Code", "Qty"],
            [topDataRowIndex]: ["r1", "a", "loose-between", "c1", 5],
            [topDataRowIndex + 1]: ["r2", "b", "", "c2", 6],
            [topDataRowIndex + 3]: ["loose-below"],
          }),
          tables: [
            {
              tableId: "left",
              headRows: { 3: ["lft:1", "lft:2"] },
              startColumnIndex: leftStart,
              endColumnIndex: leftStart + 2,
              endRowIndex: topDataRowIndex + 2,
            },
            {
              tableId: "right",
              headRows: { 3: ["rgt:1", "rgt:2"] },
              startColumnIndex: rightStart,
              endColumnIndex: rightStart + 2,
              endRowIndex: topDataRowIndex + 2,
            },
          ].filter(({ tableId }) => tableIds.includes(tableId)),
        },
      ],
    });
  }

  it("keeps each Table's cells in its own state, counted from its own first column", () => {
    stubSideBySideTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchSheetUsedGrid(scratchGid);

    expect(raw.sheet(scratchGid).tableIds()).toEqual(["left", "right"]);
    expect(raw.table("left").row(0).activeValueArr).toEqual(["r1", "a"]);
    expect(raw.table("right").row(0).activeValueArr).toEqual(["c1", 5]);
    expect(raw.table("right").row(1).valueOrEmpty(1)).toBe(6);
    expect(raw.table("left").meta.colIdRow.activeValueArr).toEqual([
      "lft:1",
      "lft:2",
    ]);
    expect(raw.table("right").meta.colIdRow.activeValueArr).toEqual([
      "rgt:1",
      "rgt:2",
    ]);
  });

  it("keeps each Table's queued writes apart, landing in its own columns", () => {
    const { grid } = stubSideBySideTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchSheetUsedGrid(scratchGid);
    raw.table("left").column(1).updateAllCells({ value: "x" });
    raw.table("right").column(1).updateAllCells({ value: 0 });
    raw.table("right").row(0).updateValue(0, "c9");
    raw.batchUpdateGSheets();

    expect(
      grid.sheet(scratchGid).values({
        startRowIndex: topDataRowIndex,
        endRowIndex: topDataRowIndex + 2,
      }),
    ).toEqual([
      ["r1", "x", "loose-between", "c9", 0],
      ["r2", "x", "", "c2", 0],
    ]);
  });

  it("marks every Table on the sheet stale once a row delete there is sent", () => {
    stubSideBySideTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchSheetUsedGrid(scratchGid);
    raw.table("left").row(0).delete();
    raw.batchUpdateGSheets();

    expect(() => raw.table("left").dataRowCount).toThrow(/stale/);
    expect(() => raw.table("right").dataRowCount).toThrow(/stale/);
  });

  it("refuses a re-fetch that drops a Table still holding queued writes", () => {
    stubSideBySideTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchSheetUsedGrid(scratchGid);
    raw.table("right").row(0).updateValue(0, "c9");
    stubSideBySideTables(["left"]);
    // The same run now reads a spreadsheet where the right Table is gone.
    raw.spreadsheetRawProps.spreadsheetStateRaw.rawSource =
      installedRawSource();

    expect(() => raw.fetchSheetUsedGrid(scratchGid)).toThrow(
      /Table right is no longer on .*queued writes/,
    );
  });

  it("refuses a re-fetch down to one known Table while writes queued through the sheet wait", () => {
    stubSideBySideTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchSheetUsedGrid(scratchGid);
    raw
      .sheet(scratchGid)
      .requestSortGSheet({ colIdxToSortBy: 0, sortOrder: "DESCENDING" });
    stubSideBySideTables(["left"]);
    raw.spreadsheetRawProps.spreadsheetStateRaw.rawSource =
      installedRawSource();

    expect(() => raw.fetchSheetUsedGrid(scratchGid)).toThrow(
      /queued through .* while it had several Tables/,
    );
  });

  it("drops a loose cell beside, above or below the Tables", () => {
    stubSideBySideTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchSheetUsedGrid(scratchGid);

    const headAndBodyRowIndexes = [-4, -3, -2, -1, 0, 1];
    expect(raw.table("left").activeRowIndexes).toEqual(headAndBodyRowIndexes);
    expect(raw.table("right").activeRowIndexes).toEqual(headAndBodyRowIndexes);
    expect(raw.table("left").row(0).cell(2).isActive).toBe(false);
  });
});
