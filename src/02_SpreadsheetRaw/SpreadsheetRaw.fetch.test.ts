import { Val } from "@byronbroughten/utils/val";
import { describe, expect, it } from "vitest";

import { installedRawSource } from "../00_Source/RawSource/RawSource";
import { getTableTraitByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
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
  itemTableId,
  layoutBodyRows,
  layoutGid,
  layoutSheet,
  layoutTableId,
  layoutTableNames,
  lightGreen,
  managedLayoutTableNames,
  misplacedTableSheet,
  ownColumnId,
  placedTableSheet,
  recordedGridRanges,
  scratchGid,
  scratchTableId,
  startTableColIndex,
  tableEndRowIndex,
  tableHeaderRowIndex,
  tableId111,
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

  it("keeps every Table of a sheet holding several on the unfiltered census", () => {
    stubSheetsService({ sheets: [layoutSheet()] });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();

    expect(raw.sheet(layoutGid).tableIds).toEqual(
      layoutTableNames.map(layoutTableId),
    );
    expect(
      layoutTableNames.map(
        (tableName) => raw.table(layoutTableId(tableName)).hasFetchedProperties,
      ),
    ).toEqual([true, true, true]);
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
  it("leaves a sheet the config does not know alone when its header zone finds no Table", () => {
    stubSheetsService({
      sheets: [{ sheetId: scratchGid, title: "Byron's Scratch Sheet" }],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(scratchGid).gatherFetchHeaderZone();

    expect(() => raw.fetchAllGathered()).not.toThrow();
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
    raw.sheet(111).gatherFetchHeaderZone();

    expect(() => raw.fetchAllGathered()).not.toThrow();
  });

  it("does not throw for a config-known sheet whose Table starts where the layout requires", () => {
    stubSheetsService({
      sheets: [placedTableSheet({ sheetId: itemGid, title: "Item" })],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(itemGid).table(itemTableId).gatherFetchProperties();

    expect(() => raw.fetchAllGathered()).not.toThrow();
  });

  it("fetches a recorded Table outside the header zone without judging it", () => {
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
    raw.sheet(itemGid).table(itemTableId).gatherFetchProperties();
    raw.fetchAllGathered();

    expect(raw.table(itemTableId).startRowIndex).toBe(tableHeaderRowIndex - 1);
    expect(raw.usedTableIds).toEqual([itemTableId]);
  });

  it("accepts a Table moved right within the header zone", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: itemGid,
          title: "Item",
          rows: buildGridRows({
            [colIdRowIndex]: ["", ownColumnId(itemGid)],
            [tableHeaderRowIndex]: ["", "ID"],
          }),
          table: {
            endRowIndex: tableEndRowIndex,
            startColumnIndex: startTableColIndex + 1,
          },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(itemGid).table(itemTableId).gatherFetchProperties();

    expect(() => raw.fetchAllGathered()).not.toThrow();
    expect(raw.table(itemTableId).startColumnIndex).toBe(
      startTableColIndex + 1,
    );
  });

  it("fetches a Table met with only its header, leaving the refusal to its position reads", () => {
    stubSheetsService({
      sheets: [
        {
          ...placedTableSheet({ sheetId: itemGid, title: "Item" }),
          table: { endRowIndex: topDataRowIndex, name: "Items" },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(itemGid).table(itemTableId).gatherFetchProperties();

    expect(() => raw.fetchAllGathered()).not.toThrow();
    expect(thrownMessage(() => raw.table(itemTableId).dataRowCount)).toMatch(
      /^Table "Items" on "Item" \(gid \d+\) has only its header: add a row below it holding its formulas\.$/,
    );
  });

  it("leaves a header-only Table on a sheet the config does not know to the code that reaches it", () => {
    stubSheetsService({
      sheets: [
        {
          ...placedTableSheet({ sheetId: scratchGid, title: "Scratch" }),
          table: { endRowIndex: topDataRowIndex },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(scratchGid).gatherFetchHeaderZone();

    expect(() => raw.fetchAllGathered()).not.toThrow();
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
    raw.sheet(scratchGid).gatherFetchHeaderZone();

    expect(() => raw.fetchAllGathered()).not.toThrow();
  });

  it("checks a well-placed Table's column IDs in the one round trip", () => {
    const { getByDataFilterCalls, getCalls } = stubSheetsService({
      sheets: [placedTableSheet({ sheetId: itemGid, title: "Item" })],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(itemGid).table(itemTableId).gatherFetchProperties();
    raw.table(itemTableId).columnResolver.gatherFetchColumnIds();
    raw.fetchAllGathered();

    expect(getByDataFilterCalls).toHaveLength(1);
    expect(getCalls).toEqual([]);
  });

  it("accepts a blank beside the Table's own column IDs", () => {
    stubSheetsService({
      sheets: [
        {
          ...placedTableSheet({ sheetId: itemGid, title: "Item" }),
          rows: buildGridRows({
            [colIdRowIndex]: [ownColumnId(itemGid), ""],
            [tableHeaderRowIndex]: ["ID", "Name"],
          }),
          table: { endRowIndex: tableEndRowIndex, endColumnIndex: 2 },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(itemGid).table(itemTableId).gatherFetchProperties();
    raw.table(itemTableId).columnResolver.gatherFetchColumnIds();

    expect(() => raw.fetchAllGathered()).not.toThrow();
  });

  it("lists an unmanaged Table beside the recorded one on the sheet's Tables", () => {
    const recordedTableId = getTableTraitByName("item", "tableId");
    const { table, ...sheet } = placedTableSheet({
      sheetId: itemGid,
      title: "Item",
    });
    stubSheetsService({
      sheets: [
        {
          ...sheet,
          tables: [
            {
              tableId: "unmanaged",
              startColumnIndex: 20,
              endColumnIndex: 22,
              endRowIndex: tableEndRowIndex,
            },
            { ...Val.assert(table, "placed Table"), tableId: recordedTableId },
          ],
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();

    expect(raw.sheet(itemGid).tableIds.sort()).toEqual(
      ["unmanaged", recordedTableId].sort(),
    );
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
    raw.sheet(scratchGid).gatherFetchHeaderZone();

    expect(() => raw.fetchAllGathered()).not.toThrow();
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

  it("aims the header zone at the sheet's top rows, across every column", () => {
    const { getByDataFilterCalls } = stubSheetsService({
      sheets: [placedTableSheet({ sheetId: itemGid, title: "Item" })],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(itemGid).table(itemTableId).gatherFetchProperties();
    raw.fetchAllGathered();

    expect(recordedGridRanges(getByDataFilterCalls)).toEqual([
      {
        sheetId: itemGid,
        startRowIndex: 0,
        endRowIndex: tableHeaderRowIndex + 1,
      },
    ]);
  });

  it("lets the column ID row ride the header zone before the Table is known", () => {
    const { getByDataFilterCalls } = stubSheetsService({
      sheets: [placedTableSheet({ sheetId: itemGid, title: "Item" })],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(itemGid).table(itemTableId).columnResolver.gatherFetchColumnIds();
    raw.fetchAllGathered();

    expect(recordedGridRanges(getByDataFilterCalls)).toEqual([
      {
        sheetId: itemGid,
        startRowIndex: 0,
        endRowIndex: tableHeaderRowIndex + 1,
      },
    ]);
  });

  it("refuses to aim a gathered write at a Table not yet fetched, since no position is recorded", () => {
    stubSheetsService({
      sheets: [placedTableSheet({ sheetId: itemGid, title: "Item" })],
    });

    const raw = SpreadsheetRaw.init();

    expect(() =>
      raw.sheet(itemGid).table(itemTableId).originAtGathering(),
    ).toThrowError(/Table is unknown for sheetGid/);
  });

  it("refuses a full-row fetch before the sheet has a Table in state", () => {
    stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records" }],
    });

    const raw = SpreadsheetRaw.init();

    expect(() =>
      raw.sheet(111).table(tableId111).topRow.gatherFetchFull(),
    ).toThrowError(/Table is unknown for sheetGid 111/);
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
    raw.table(scratchTableId).topRow.gatherFetchFull();
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
    raw.table(tableId111).column(1).gatherFetchFull();
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
    expect(raw.table(tableId111).column(1).valueArrOrEmpty).toEqual(["a", "b"]);
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
    const sheet = raw.table(tableId111);
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
    raw.table(scratchTableId).columnResolver.gatherFetchColumnIds();
    raw.table(scratchTableId).row(0).gatherFetchFull();
    raw.table(scratchTableId).row(1).gatherFetchFull();
    raw.fetchAllGathered();
    return raw;
  }

  it("reads body row 0 from the row just below the header, and column 0 from the Table's first column", () => {
    stubLowerTable();
    const raw = fetchedLowerRaw();
    const sheet = raw.table(scratchTableId);

    expect(sheet.rowIndexesFull).toEqual([0, 1]);
    expect(sheet.row(0).valueOrEmpty(0)).toBe("r:low:1");
    expect(sheet.row(1).valueOrEmpty(1)).toBe("second");
    expect(sheet.profile.columnIds).toEqual(["c:low:aaa", "c:low:bbb"]);
    expect(raw.table(scratchTableId).headRow("header").valueOrEmpty(1)).toBe(
      "Name",
    );
  });

  it("updates and appends at the Table's grid rows, leaving the loose cell beside it alone", () => {
    const { grid } = stubLowerTable();
    const raw = fetchedLowerRaw();
    const sheet = raw.table(scratchTableId);

    sheet.row(1).cell(1).updateValue("changed");
    sheet.appendDataRow().cell(0).updateValue("r:low:3");
    sheet.column(1).updateWorkingCells({ backgroundColor: lightGreen });
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

    expect(() => raw.table(scratchTableId).row(5).valueOrEmpty(0)).toThrowError(
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

    expect(raw.table("left").tableIds()).toEqual(["left", "right"]);
    expect(raw.table("left").row(0).workingValueArr).toEqual(["r1", "a"]);
    expect(raw.table("right").row(0).workingValueArr).toEqual(["c1", 5]);
    expect(raw.table("right").row(1).valueOrEmpty(1)).toBe(6);
    expect(raw.table("left").headRow("columnId").workingValueArr).toEqual([
      "lft:1",
      "lft:2",
    ]);
    expect(raw.table("right").headRow("columnId").workingValueArr).toEqual([
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

  it("drops a loose cell beside, above or below the Tables", () => {
    stubSideBySideTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchSheetUsedGrid(scratchGid);

    const headAndBodyRowIndexes = [-4, -3, -2, -1, 0, 1];
    expect(raw.table("left").workingRowIndexesWithHead).toEqual(
      headAndBodyRowIndexes,
    );
    expect(raw.table("right").workingRowIndexesWithHead).toEqual(
      headAndBodyRowIndexes,
    );
    expect(raw.table("left").row(0).cell(2).inWorking).toBe(false);
  });
});

describe("several managed Tables on one sheet", () => {
  const leftId = layoutTableId("layoutLeft");
  const rightId = layoutTableId("layoutRight");

  // Moves layoutRight one column right under a run that has already fetched.
  function stubMovedRight(raw: SpreadsheetRaw): void {
    stubSheetsService({
      isEveryTableInFilteredFetch: true,
      sheets: [layoutSheet({ layoutRight: { startColumnIndex: 4 } })],
    });
    raw.spreadsheetRawProps.spreadsheetStateRaw.rawSource =
      installedRawSource();
  }
  function gatherEveryZone(raw: SpreadsheetRaw): void {
    managedLayoutTableNames.forEach((tableName) => {
      const table = raw.sheet(layoutGid).table(layoutTableId(tableName));
      table.gatherFetchProperties();
      table.columnResolver.gatherFetchColumnIds();
    });
  }

  it("finds every managed Table on the sheet through one header zone", () => {
    const { getByDataFilterCalls } = stubSheetsService({
      sheets: [layoutSheet()],
    });

    const raw = SpreadsheetRaw.init();
    gatherEveryZone(raw);
    raw.fetchAllGathered();

    expect(
      managedLayoutTableNames.map((tableName) => {
        const origin = raw.table(layoutTableId(tableName)).tableOrigin();
        return [origin.headerRowIndex, origin.startColIndex];
      }),
    ).toEqual([
      [3, 0],
      [3, 3],
    ]);
    expect(recordedGridRanges(getByDataFilterCalls)).toEqual([
      { sheetId: layoutGid, startRowIndex: 0, endRowIndex: 4 },
    ]);
  });

  it("reads each Table's body from its own rows and columns", () => {
    stubSheetsService({ sheets: [layoutSheet()] });

    const raw = SpreadsheetRaw.init();
    gatherEveryZone(raw);
    raw.fetchAllGathered();
    managedLayoutTableNames.forEach((tableName) => {
      const table = raw.table(layoutTableId(tableName));
      layoutBodyRows[tableName].forEach((_, rowIndex) =>
        table.row(rowIndex).gatherFetchFull(),
      );
    });
    raw.fetchAllGathered();

    expect(
      managedLayoutTableNames.map((tableName) =>
        layoutBodyRows[tableName].map((_, rowIndex) =>
          [0, 1].map((colIndex) =>
            raw
              .table(layoutTableId(tableName))
              .row(rowIndex)
              .valueOrEmpty(colIndex),
          ),
        ),
      ),
    ).toEqual(
      managedLayoutTableNames.map((tableName) => layoutBodyRows[tableName]),
    );
  });

  it("reads on past a neighbour moved within the header zone mid-run", () => {
    stubSheetsService({
      isEveryTableInFilteredFetch: true,
      sheets: [layoutSheet()],
    });
    const raw = SpreadsheetRaw.init();
    raw.sheet(layoutGid).table(leftId).gatherFetchProperties();
    raw.fetchAllGathered();
    stubMovedRight(raw);
    raw.table(leftId).row(0).gatherFetchFull();

    expect(() => raw.fetchAllGathered()).not.toThrow();
    expect(raw.table(rightId).startColumnIndex).toBe(4);
  });

  it("reads on past a neighbour the census listed, once it moves within the header zone", () => {
    stubSheetsService({
      isEveryTableInFilteredFetch: true,
      sheets: [layoutSheet()],
    });
    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(leftId).row(0).gatherFetchFull();
    raw.fetchAllGathered();
    stubMovedRight(raw);
    raw.table(leftId).row(0).gatherFetchFull();

    expect(() => raw.fetchAllGathered()).not.toThrow();
  });

  it("reads a Table moved right within the header zone without regenerating", () => {
    stubSheetsService({
      sheets: [layoutSheet({ layoutRight: { startColumnIndex: 4 } })],
    });

    const raw = SpreadsheetRaw.init();
    gatherEveryZone(raw);
    raw.fetchAllGathered();

    expect(raw.table(rightId).tableOrigin().startColIndex).toBe(4);
  });

  it("reads one Table while a neighbour it doesn't use is missing", () => {
    const sheet = layoutSheet();
    stubSheetsService({
      sheets: [
        {
          ...sheet,
          tables: sheet.tables?.filter((table) => table.tableId !== rightId),
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(layoutGid).table(leftId).gatherFetchProperties();
    raw.table(leftId).columnResolver.gatherFetchColumnIds();

    expect(() => raw.fetchAllGathered()).not.toThrow();
    expect(raw.table(leftId).columnResolver.hasFetchedColumnIds).toBe(true);
  });

  it("reads past a missing neighbour when a fetch names only the shared sheet", () => {
    const sheet = layoutSheet();
    stubSheetsService({
      sheets: [
        {
          ...sheet,
          tables: sheet.tables?.filter((table) => table.tableId !== rightId),
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(layoutGid).table(leftId).gatherFetchProperties();
    raw.table(leftId).columnResolver.gatherFetchColumnIds();

    expect(() => raw.fetchAllGathered()).not.toThrow();
  });

  it("reads one Table while a neighbour it doesn't use holds another Table's column IDs", () => {
    stubSheetsService({
      sheets: [
        layoutSheet({
          layoutRight: {
            headRows: { 3: ["c:lyl:entry", "c:lyl:amount"] },
          },
        }),
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(layoutGid).table(leftId).gatherFetchProperties();
    raw.table(leftId).columnResolver.gatherFetchColumnIds();

    expect(() => raw.fetchAllGathered()).not.toThrow();
  });

  it("keeps a Table a filtered fetch did not return, with what was queued on it", () => {
    const { grid } = stubSheetsService({
      sheets: [layoutSheet()],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(rightId).row(0).updateValue(1, 11);
    raw.table(leftId).row(0).gatherFetchFull();
    raw.fetchAllGathered();
    raw.batchUpdateGSheets();

    expect(raw.table(rightId).hasFetchedProperties).toBe(true);
    expect(grid.sheet(layoutGid).cell(4, 4)).toBe(11);
  });
});
