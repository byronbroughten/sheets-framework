import { describe, expect, it } from "vitest";

import type { CellValue } from "../00_Source/CellValues/cellValues";
import { googleRawRequest } from "../00_Source/GoogleSheets/GoogleSheetsAPI";
import type { AddTableOperation } from "../00_Source/RawSource/RawSource";
import type { RgbColor } from "../00_Source/RawSource/RgbColor";
import { getSheetTraitByName } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import { ssConfigGet } from "../01_SpreadsheetSchema/spreadsheetConfigTypes";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  buildGridRows,
  type FakeCell,
  type FakeSheetProperties,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { assertType, type IsExactly } from "../testSupport/typeAssertions";
import type { CellRaw } from "./CellRaw";
import type { RowCommonRaw } from "./ClassBases/RowCommonRaw";
import type { CellStateRaw, RowCellChange } from "./ClassTypes/StateRaw";
import { ColumnMetaRaw } from "./ColumnMetaRaw";
import { ColumnRaw } from "./ColumnRaw";
import { RowRaw } from "./RowRaw";
import { SheetMetaRaw } from "./SheetMetaRaw";
import { SheetRaw } from "./SheetRaw";
import { SpreadsheetRaw } from "./SpreadsheetRaw";
import { UniformRowRaw } from "./UniformRowRaw";

const lightGreen = { red: 0.851, green: 0.918, blue: 0.827 };

const itemGid = getSheetTraitByName("item", "sheetGid");
const logGid = getSheetTraitByName("log", "sheetGid");
const tableHeaderRowIndex = ssConfigGet("tableHeaderRowIndexBase0");
const colIdRowIndex = ssConfigGet("columnIdRowIdxBase0");
const startTableColIndex = ssConfigGet("startTableColIndexBase0");
const topDataRowIndex = tableHeaderRowIndex + 1;
const scratchGid = 999999;
const tableEndRowIndex = tableHeaderRowIndex + 3;

function placedTableSheet(sheet: {
  sheetId: number;
  title: string;
}): FakeSheetProperties {
  return {
    ...sheet,
    rows: buildGridRows({ [tableHeaderRowIndex]: ["ID"] }),
    table: { endRowIndex: tableEndRowIndex },
  };
}

function misplacedTableSheet({
  startRowIndex = tableHeaderRowIndex,
  startColumnIndex = startTableColIndex,
  ...sheet
}: {
  sheetId: number;
  title: string;
  startRowIndex?: number;
  startColumnIndex?: number;
}): FakeSheetProperties {
  return {
    ...placedTableSheet(sheet),
    table: {
      endRowIndex: tableEndRowIndex,
      startRowIndex,
      startColumnIndex,
    },
  };
}

function extraTablesSheet(sheet: {
  sheetId: number;
  title: string;
}): FakeSheetProperties {
  return {
    ...placedTableSheet(sheet),
    extraTables: [
      {
        startRowIndex: tableHeaderRowIndex + 10,
        endRowIndex: tableHeaderRowIndex + 12,
      },
    ],
  };
}

function recordedGridRanges(calls: object[]): unknown[] {
  const resource = calls[0] as {
    dataFilters: { gridRange: unknown }[];
  };
  return resource.dataFilters.map((filter) => filter.gridRange);
}

function thrownMessage(fn: () => void): string {
  try {
    fn();
  } catch (error) {
    return (error as Error).message;
  }
  throw new Error("Expected the call to throw, but it did not.");
}

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
    raw.sheet(111).gatherFetchProperties(startTableColIndex);
    raw.sheetMeta(111).gatherFetchColumnIdsInit(startTableColIndex);
    raw.sheet(222).gatherFetchProperties(startTableColIndex);
    raw.sheetMeta(222).gatherFetchColumnIdsInit(startTableColIndex);

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
    raw.sheet(111).gatherFetchProperties(startTableColIndex);

    expect(() => raw.fetchAllGathered()).not.toThrow();
  });

  it("does not throw for a config-known sheet whose Table starts where the layout requires", () => {
    stubSheetsService({
      sheets: [placedTableSheet({ sheetId: itemGid, title: "Item" })],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(itemGid).gatherFetchProperties(startTableColIndex);

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
    raw.sheet(itemGid).gatherFetchProperties(startTableColIndex);

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
    raw.sheet(itemGid).gatherFetchProperties(startTableColIndex);

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
    raw.sheet(scratchGid).gatherFetchProperties(startTableColIndex);

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
    raw.sheet(itemGid).gatherFetchProperties(startTableColIndex);

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
    raw.sheet(itemGid).gatherFetchProperties(startTableColIndex);
    raw.sheetMeta(itemGid).gatherFetchColumnIdsInit(startTableColIndex);

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
    raw.sheet(itemGid).gatherFetchProperties(startTableColIndex);

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
    raw.sheet(itemGid).gatherFetchProperties(startTableColIndex);
    raw.sheet(logGid).gatherFetchProperties(startTableColIndex);

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
    raw.sheet(scratchGid).gatherFetchProperties(startTableColIndex);

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
    raw.sheet(itemGid).gatherFetchProperties(startTableColIndex);
    raw.sheet(logGid).gatherFetchProperties(startTableColIndex);
    raw.sheetMeta(logGid).gatherFetchColumnIdsInit(startTableColIndex);

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
    raw.sheet(itemGid).gatherFetchProperties(startTableColIndex);
    raw.sheet(logGid).gatherFetchProperties(startTableColIndex);

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
    raw.sheet(itemGid).gatherFetchProperties(startTableColIndex);
    raw.sheetMeta(itemGid).gatherFetchColumnIdsInit(startTableColIndex);

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
    raw.sheet(itemGid).gatherFetchProperties(startTableColIndex);
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
    raw.sheetMeta(itemGid).gatherFetchColumnIdsInit(startTableColIndex);
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
      /Active table is null for sheetGid 111/,
    );
  });

  it("aims a full-row fetch at the live Table start after properties, not the layout constant", () => {
    const liveStart = startTableColIndex + 1;
    const { getByDataFilterCalls } = stubSheetsService({
      sheets: [
        misplacedTableSheet({
          sheetId: scratchGid,
          title: "Byron's Scratch Sheet",
          startColumnIndex: liveStart,
        }),
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
      },
    ]);
  });
});

describe("ColumnMetaRaw active facts", () => {
  const tableEndRow = topDataRowIndex + 1;

  function stubSheetWithTopDataRow(
    topDataRow: FakeCell[],
    absence?: "rowsWithNoGridData" | "rowsWithNoGridBlock",
  ) {
    stubSheetsService({
      sheets: [
        {
          sheetId: itemGid,
          title: "Item",
          rows: buildGridRows({
            0: ["c:itm:aaa", "c:itm:bbb"],
            [tableHeaderRowIndex]: ["Purchase Price", "Notes"],
            [topDataRowIndex]: topDataRow,
          }),
          ...(absence ? { [absence]: [topDataRowIndex] } : {}),
          table: { endRowIndex: tableEndRow },
        },
      ],
    });
  }

  function fetchedItemColumnMeta(colIndex: number): ColumnMetaRaw {
    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(itemGid).topRow.gatherFetchFull();
    raw.fetchAllGathered(true);
    return raw.sheetMeta(itemGid).column(colIndex);
  }

  function expectBlankFacts(column: ColumnMetaRaw): void {
    expect(column.activeIsFormula).toBe(false);
    expect(column.activeNumberFormatType).toBeUndefined();
    expect(column.activeTopValue).toBe("");
  }

  it("reports blank facts for a top data row returned without any cell data", () => {
    stubSheetWithTopDataRow([], "rowsWithNoGridData");

    expectBlankFacts(fetchedItemColumnMeta(0));
  });

  it("reports blank facts for a top data row returned as no grid block at all", () => {
    stubSheetWithTopDataRow([], "rowsWithNoGridBlock");

    expectBlankFacts(fetchedItemColumnMeta(0));
  });

  it("reports the same facts an empty cell inside a returned row produces", () => {
    stubSheetWithTopDataRow([null, "a note"]);

    expectBlankFacts(fetchedItemColumnMeta(0));
  });

  it("keeps the facts the payload supplied rather than seeding over them", () => {
    stubSheetWithTopDataRow([
      { value: 42, isFormula: true, numberFormatType: "CURRENCY" },
    ]);

    const column = fetchedItemColumnMeta(0);

    expect(column.activeIsFormula).toBe(true);
    expect(column.activeNumberFormatType).toBe("CURRENCY");
    expect(column.activeTopValue).toBe(42);
  });

  it("reports blank facts for a full-column fetch of a wholly blank column", () => {
    stubSheetWithTopDataRow([], "rowsWithNoGridData");

    const raw = SpreadsheetRaw.init();
    raw.sheet(itemGid).gatherFetchProperties(startTableColIndex);
    raw.fetchAllGathered();
    raw.sheet(itemGid).column(1).gatherFetchFull();
    raw.fetchAllGathered(true);

    expectBlankFacts(raw.sheetMeta(itemGid).column(1));
  });

  it("reads a specifically fetched cell omitted from the payload as empty, not unfetched", () => {
    stubSheetWithTopDataRow([], "rowsWithNoGridData");

    const raw = SpreadsheetRaw.init();
    raw.sheet(itemGid).gatherFetchProperties(startTableColIndex);
    raw.fetchAllGathered();
    const cell = raw.sheet(itemGid).row(topDataRowIndex).cell(0);
    cell.gatherFetchRange();
    raw.fetchAllGathered();

    expect(cell.valueOrEmpty()).toBe("");
  });

  it("throws naming the sheet and the missing fetch for a column nothing fetched", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: itemGid,
          title: "Item",
          rows: buildGridRows({ [tableHeaderRowIndex]: ["Purchase Price"] }),
          table: { endRowIndex: tableEndRow },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(itemGid).gatherFetchProperties(startTableColIndex);
    raw.fetchAllGathered();

    const message = thrownMessage(
      () => raw.sheetMeta(itemGid).column(0).activeIsFormula,
    );
    expect(message).toContain(`"Item" (gid ${itemGid})`);
    expect(message).toMatch(/top data row/);
  });

  it("writes no facts for a grid column outside the table", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: itemGid,
          title: "Item",
          rows: buildGridRows({
            0: ["c:itm:aaa"],
            [tableHeaderRowIndex]: ["Purchase Price"],
            [topDataRowIndex]: [100000, "outside the table"],
          }),
          table: { endRowIndex: tableEndRow, endColumnIndex: 1 },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(itemGid).topRow.gatherFetchFull();
    raw.fetchAllGathered(true);

    expect(raw.sheetMeta(itemGid).column(0).activeTopValue).toBe(100000);
    expect(() => raw.sheetMeta(itemGid).column(1).activeTopValue).toThrowError(
      /No active facts/,
    );
  });
});

describe("SpreadsheetRaw.batchUpdateGSheets", () => {
  it("sends exactly the sort request gathered for a sheet-level sort change", () => {
    const { batchUpdateCalls } = stubSheetsService();

    const raw = SpreadsheetRaw.init();
    raw.sheet(111).requestSortGSheet({
      colIdxToSortBy: 2,
      sortOrder: "ASCENDING",
    });
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls).toEqual([
      {
        requests: [
          {
            sortRange: {
              range: { sheetId: 111, startRowIndex: 4, startColumnIndex: 0 },
              sortSpecs: [{ dimensionIndex: 2, sortOrder: "ASCENDING" }],
            },
          },
        ],
      },
    ]);
  });

  it("sends no request when there is nothing to save", () => {
    const { batchUpdateCalls } = stubSheetsService();

    SpreadsheetRaw.init().batchUpdateGSheets();

    expect(batchUpdateCalls).toEqual([]);
  });

  it("sends one appendCells whose rows array is the full append, so a Sheets table grows by every row rather than by one", () => {
    const { batchUpdateCalls } = stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).appendDataRow();
    raw.sheet(111).appendDataRow();
    raw.sheet(111).appendDataRow();
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        appendCells: {
          sheetId: 111,
          tableId: "fake-table-111",
          rows: [{}, {}, {}],
          fields: "userEnteredValue",
        },
      },
    ]);
  });

  it("keeps a second sheet's append as its own request rather than folding it into the first table's", () => {
    const { batchUpdateCalls } = stubSheetsService({
      sheets: [
        { sheetId: 111, title: "Records", table: { endRowIndex: 11 } },
        { sheetId: 222, title: "Entries", table: { endRowIndex: 6 } },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).appendDataRow();
    raw.sheet(222).appendDataRow();
    raw.sheet(111).appendDataRow();
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        appendCells: {
          sheetId: 111,
          tableId: "fake-table-111",
          rows: [{}, {}],
          fields: "userEnteredValue",
        },
      },
      {
        appendCells: {
          sheetId: 222,
          tableId: "fake-table-222",
          rows: [{}],
          fields: "userEnteredValue",
        },
      },
    ]);
  });

  it("still gathers an append queued after a deletion on the same sheet, since row indexes only shift once the deletes are sent", () => {
    const { batchUpdateCalls } = stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).row(5).delete();
    raw.sheet(111).appendDataRow();

    expect(() => raw.batchUpdateGSheets()).not.toThrow();
    expect(batchUpdateCalls[0]?.requests?.[0]).toEqual({
      appendCells: {
        sheetId: 111,
        tableId: "fake-table-111",
        rows: [{}],
        fields: "userEnteredValue",
      },
    });
    expect(raw.sheet(111).rowIndexesAreStale).toBe(true);
  });

  const staleRowIndexes = "Row indexes are stale for sheetGid 111.";

  function sheetAfterFlushedDataRowDelete(fetchKeptRow = false) {
    stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({
            4: ["kept"],
            5: ["deleted"],
            6: ["later"],
          }),
          table: {
            endRowIndex: 11,
            endColumnIndex: 3,
            columnTypes: { 0: "TEXT" },
            columnValidationValues: { 0: ["=valueConfig[Notes]"] },
            columnValidationConditionTypes: { 0: "BOOLEAN" },
          },
        },
      ],
    });
    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    if (fetchKeptRow) {
      raw.sheet(111).row(4).gatherFetchFull();
      raw.fetchAllGathered();
    }
    raw.sheet(111).row(5).delete();
    raw.batchUpdateGSheets();
    return raw;
  }

  it("still reads table column properties after a flushed row delete, while the table end throws", () => {
    const raw = sheetAfterFlushedDataRowDelete();

    const table = raw.sheet(111).activeTable;
    expect(table.tableId).toBe("fake-table-111");
    expect(table.startRowIndex).toBe(tableHeaderRowIndex);
    expect(table.startColumnIndex).toBe(startTableColIndex);
    expect(table.endColumnIndex).toBeGreaterThan(startTableColIndex);
    const columnMeta = raw.sheet(111).meta.column(0);
    expect(columnMeta.activeColumnType).toBe("TEXT");
    expect(columnMeta.valueValidationStrings).toEqual(["=valueConfig[Notes]"]);
    expect(columnMeta.validationConditionType).toBe("BOOLEAN");
    expect(raw.sheet(111).isTableColIndex(startTableColIndex)).toBe(true);
    expect(() => table.endRowIndex).toThrow(staleRowIndexes);
    expect(() => {
      table.endRowIndex = 12;
    }).toThrow(staleRowIndexes);
  });

  it("keys declared column types by sheet column when the Table starts after column A", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          table: {
            startColumnIndex: 1,
            endRowIndex: 11,
            endColumnIndex: 4,
            columnTypes: { 1: "TEXT" },
            columnValidationValues: { 1: ["=valueConfig[Notes]"] },
            columnValidationConditionTypes: { 1: "BOOLEAN" },
          },
        },
      ],
    });
    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();

    const columnMeta = raw.sheet(111).meta.column(1);
    expect(columnMeta.activeColumnType).toBe("TEXT");
    expect(columnMeta.valueValidationStrings).toEqual(["=valueConfig[Notes]"]);
    expect(columnMeta.validationConditionType).toBe("BOOLEAN");
  });

  it("throws on per-cell value, formula, and colour writes after a flushed row delete", () => {
    const raw = sheetAfterFlushedDataRowDelete(true);
    const cell = raw.sheet(111).row(4).cell(0);

    expect(() => cell.updateValue("painted")).toThrow(staleRowIndexes);
    expect(() => cell.updateFormula("=1")).toThrow(staleRowIndexes);
    expect(() => cell.updateBackgroundColor(lightGreen)).toThrow(
      staleRowIndexes,
    );
    expect(() => cell.addCheckboxValidation()).toThrow(staleRowIndexes);
  });

  it("throws on active-row and whole-column fills after a flushed row delete", () => {
    const raw = sheetAfterFlushedDataRowDelete(true);
    const column = raw.sheet(111).column(0);

    expect(() => column.updateActiveCells({ value: "fill" })).toThrow(
      staleRowIndexes,
    );
    expect(() => column.updateActiveFormulas("=1")).toThrow(staleRowIndexes);
    expect(() => column.updateAllCells({ value: "fill" })).toThrow(
      staleRowIndexes,
    );
    expect(() => column.updateAllFormulas("=1")).toThrow(staleRowIndexes);
  });

  it("throws on a further data-row delete after a flushed row delete", () => {
    const raw = sheetAfterFlushedDataRowDelete();

    expect(() => raw.sheet(111).row(6).delete()).toThrow(staleRowIndexes);
  });

  it("still reads an already-fetched cell after a flushed row delete", () => {
    const raw = sheetAfterFlushedDataRowDelete(true);

    expect(raw.sheet(111).row(4).cell(0).valueOrEmpty()).toBe("kept");
  });

  it("leaves row indexes stale after a properties fetch that follows a flushed row delete", () => {
    const raw = sheetAfterFlushedDataRowDelete();
    raw.fetchAllSheetProperties();

    expect(raw.sheet(111).rowIndexesAreStale).toBe(true);
    expect(() => raw.sheet(111).activeTable.endRowIndex).toThrow(
      staleRowIndexes,
    );
  });

  it("clears row-index stale only when clearRowIndexStale is called, and then a write is allowed again", () => {
    const raw = sheetAfterFlushedDataRowDelete(true);
    const cell = raw.sheet(111).row(4).cell(0);
    expect(() => cell.updateValue("painted")).toThrow(staleRowIndexes);

    raw.sheet(111).clearRowIndexStale();

    expect(raw.sheet(111).rowIndexesAreStale).toBe(false);
    expect(raw.sheet(111).activeTable.endRowIndex).toBe(11);
    expect(() => cell.updateValue("painted")).not.toThrow();
  });

  it("does not mark row indexes stale when a queued delete is discarded before flush", () => {
    stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).row(5).delete();
    raw.discardQueuedChanges();

    expect(raw.sheet(111).rowIndexesAreStale).toBe(false);
    expect(() => raw.sheet(111).row(4).cell(0).updateValue("ok")).not.toThrow();
  });

  it("sends same-sheet row deletions in descending startIndex order so an earlier deletion can't shift a later one out from under it", () => {
    const { batchUpdateCalls } = stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).row(5).delete();
    raw.sheet(111).row(10).delete();
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls).toEqual([
      {
        requests: [
          {
            deleteDimension: {
              range: {
                sheetId: 111,
                dimension: "ROWS",
                startIndex: 10,
                endIndex: 11,
              },
            },
          },
          {
            deleteDimension: {
              range: {
                sheetId: 111,
                dimension: "ROWS",
                startIndex: 5,
                endIndex: 6,
              },
            },
          },
        ],
      },
    ]);
  });
});

describe("SpreadsheetRaw.gatherRawRequest", () => {
  it("sends a raw request last, after every request the framework models", () => {
    const { batchUpdateCalls } = stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.gatherRawRequest(
      googleRawRequest({ updateTable: { table: { tableId: "t" } } }),
    );
    raw.sheet(111).row(5).cell(2).updateValue("Processing...");
    raw.sheet(111).row(10).delete();
    raw.batchUpdateGSheets();

    const requests = batchUpdateCalls[0]?.requests ?? [];
    expect(requests.map((request) => Object.keys(request)[0])).toEqual([
      "updateCells",
      "deleteDimension",
      "updateTable",
    ]);
  });

  it("discards a raw request alongside every other queued change", () => {
    const { batchUpdateCalls } = stubSheetsService();

    const raw = SpreadsheetRaw.init();
    raw.gatherRawRequest(
      googleRawRequest({ updateTable: { table: { tableId: "t" } } }),
    );
    raw.discardQueuedChanges();
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls).toEqual([]);
  });
});

describe("SpreadsheetRaw add sheet and add Table", () => {
  const addSheetProps = {
    sheetId: 555,
    title: "Spreadsheet Config",
    rowCount: 20,
    columnCount: 6,
  };
  const addTableProps: Omit<AddTableOperation, "kind"> = {
    name: "spreadsheetConfig",
    range: {
      sheetId: 555,
      startRowIndex: 2,
      endRowIndex: 5,
      startColumnIndex: 1,
      endColumnIndex: 3,
    },
    columnProperties: [
      { columnIndex: 1, columnName: "Name", columnType: "TEXT" },
    ],
  };

  it("sends the add-sheet request first, the add-Table request second, and a sheet-title update after both", () => {
    const { batchUpdateCalls } = stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).updateTitle("Renamed");
    raw.gatherAddTableRequest(addTableProps);
    raw.gatherAddSheetRequest(addSheetProps);
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls).toHaveLength(1);
    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        addSheet: {
          properties: {
            sheetId: 555,
            title: "Spreadsheet Config",
            gridProperties: { rowCount: 20, columnCount: 6 },
          },
        },
      },
      {
        addTable: {
          table: {
            tableId: "spreadsheetConfig",
            name: "spreadsheetConfig",
            range: addTableProps.range,
          },
        },
      },
      {
        updateTable: {
          table: {
            tableId: "spreadsheetConfig",
            columnProperties: addTableProps.columnProperties,
          },
          fields: "columnProperties",
        },
      },
      {
        updateSheetProperties: {
          properties: { sheetId: 111, title: "Renamed" },
          fields: "title",
        },
      },
    ]);
  });

  it("puts one operation on the spreadsheet's write queue per queue method", () => {
    stubSheetsService();

    const raw = SpreadsheetRaw.init();
    raw.gatherAddSheetRequest(addSheetProps);
    raw.gatherAddTableRequest(addTableProps);

    expect(raw.updateRequests.addSheet).toEqual([
      { kind: "addSheet", ...addSheetProps },
    ]);
    expect(raw.updateRequests.addTable).toEqual([
      { kind: "addTable", ...addTableProps },
    ]);
  });

  it("empties both lists on a flush", () => {
    stubSheetsService();

    const raw = SpreadsheetRaw.init();
    raw.gatherAddSheetRequest(addSheetProps);
    raw.gatherAddTableRequest(addTableProps);
    raw.batchUpdateGSheets();

    expect(raw.updateRequests.addSheet).toEqual([]);
    expect(raw.updateRequests.addTable).toEqual([]);
  });

  const seededCell = {
    sheetId: 555,
    rowIndex: 3,
    colIndex: 1,
    value: "Example",
  };

  const checkboxRange = {
    sheetId: 555,
    startRowIndex: 3,
    endRowIndex: 4,
    startColumnIndex: 1,
    endColumnIndex: 2,
  };

  it("refuses a seeded value for a GID with no add-sheet queued, naming the GID, and queues nothing", () => {
    stubSheetsService();

    const raw = SpreadsheetRaw.init();

    expect(() => raw.gatherAddedSheetCellRequest(seededCell)).toThrow(
      "Added-sheet cell write refused: no addSheet for GID 555 is queued in this flush.",
    );
    expect(raw.updateRequests.update).toEqual([]);
  });

  it("refuses a seeded value after the create flush has sent that GID's add-sheet", () => {
    stubSheetsService();

    const raw = SpreadsheetRaw.init();
    raw.gatherAddSheetRequest(addSheetProps);
    raw.batchUpdateGSheets();

    expect(() => raw.gatherAddedSheetCellRequest(seededCell)).toThrow(
      "no addSheet for GID 555",
    );
  });

  it("seeds a formula on an added tab, and refuses one that doesn't start with =", () => {
    const { batchUpdateCalls } = stubSheetsService();
    const { value: _value, ...seededPosition } = seededCell;

    const raw = SpreadsheetRaw.init();
    raw.gatherAddSheetRequest(addSheetProps);
    expect(() =>
      raw.gatherAddedSheetCellRequest({ ...seededPosition, formula: "ROW()" }),
    ).toThrow('Formula must start with "="');
    raw.gatherAddedSheetCellRequest({ ...seededPosition, formula: "=ROW()" });
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls[0]?.requests).toContainEqual({
      pasteData: {
        coordinate: { sheetId: 555, rowIndex: 3, columnIndex: 1 },
        data: '"=ROW()"',
        delimiter: "\t",
        type: "PASTE_FORMULA",
      },
    });
  });

  it("sends the add-sheet, then the add-Table, then the seeded value's updateCells in one batch", () => {
    const { batchUpdateCalls } = stubSheetsService();

    const raw = SpreadsheetRaw.init();
    raw.gatherAddSheetRequest(addSheetProps);
    raw.gatherAddTableRequest(addTableProps);
    raw.gatherAddedSheetCellRequest(seededCell);
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls).toHaveLength(1);
    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        addSheet: {
          properties: {
            sheetId: 555,
            title: "Spreadsheet Config",
            gridProperties: { rowCount: 20, columnCount: 6 },
          },
        },
      },
      {
        addTable: {
          table: {
            tableId: "spreadsheetConfig",
            name: "spreadsheetConfig",
            range: addTableProps.range,
          },
        },
      },
      {
        updateTable: {
          table: {
            tableId: "spreadsheetConfig",
            columnProperties: addTableProps.columnProperties,
          },
          fields: "columnProperties",
        },
      },
      {
        updateCells: {
          range: {
            sheetId: 555,
            startRowIndex: 3,
            endRowIndex: 4,
            startColumnIndex: 1,
            endColumnIndex: 2,
          },
          rows: [
            { values: [{ userEnteredValue: { stringValue: "Example" } }] },
          ],
          fields: "userEnteredValue",
        },
      },
    ]);
  });

  it("refuses a checkbox validation for a GID with no add-sheet queued, naming the GID, and queues nothing", () => {
    stubSheetsService();

    const raw = SpreadsheetRaw.init();

    expect(() =>
      raw.gatherAddedSheetCheckboxValidationRequest(checkboxRange),
    ).toThrowError(
      "Added-sheet checkbox validation refused: no addSheet for GID 555 is queued in this flush.",
    );
    expect(raw.updateRequests.addCheckboxValidation).toEqual([]);
  });

  it("sends a checkbox validation after the add-sheet, the add-Table and the seeded value, whatever the gather order", () => {
    const { batchUpdateCalls } = stubSheetsService();

    const raw = SpreadsheetRaw.init();
    raw.gatherAddSheetRequest(addSheetProps);
    raw.gatherAddedSheetCheckboxValidationRequest(checkboxRange);
    raw.gatherAddTableRequest(addTableProps);
    raw.gatherAddedSheetCellRequest({ ...seededCell, value: false });
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls).toHaveLength(1);
    const requests = batchUpdateCalls[0]?.requests ?? [];
    expect(requests.map((request) => Object.keys(request))).toEqual([
      ["addSheet"],
      ["addTable"],
      ["updateTable"],
      ["updateCells"],
      ["setDataValidation"],
    ]);
    expect(requests[4]).toEqual({
      setDataValidation: {
        range: checkboxRange,
        rule: { condition: { type: "BOOLEAN" } },
      },
    });
    expect(raw.updateRequests.addCheckboxValidation).toEqual([]);
  });

  it("drops a queued seeded value on discardQueuedChanges", () => {
    const { batchUpdateCalls } = stubSheetsService();

    const raw = SpreadsheetRaw.init();
    raw.gatherAddSheetRequest(addSheetProps);
    raw.gatherAddedSheetCellRequest(seededCell);
    raw.discardQueuedChanges();
    raw.batchUpdateGSheets();

    expect(raw.updateRequests.update).toEqual([]);
    expect(batchUpdateCalls).toEqual([]);
  });
});

describe("RowRaw.delete", () => {
  function stubSheetWithDataRows(dataRowCount: number) {
    return stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          table: { endRowIndex: topDataRowIndex + dataRowCount },
        },
      ],
    });
  }

  it("refuses to delete the only data row, since a new row copies its formulas from the rows already there", () => {
    stubSheetWithDataRows(1);

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();

    expect(() => raw.sheet(111).topRow.delete()).toThrowError(
      /last data row.*may never be left with none/,
    );
  });

  it("refuses the delete that would take the last of several to zero", () => {
    stubSheetWithDataRows(3);

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).row(4).delete();
    raw.sheet(111).row(5).delete();

    expect(() => raw.sheet(111).row(6).delete()).toThrowError(
      /last data row.*may never be left with none/,
    );
  });

  it("still emits a row deletion when other data rows survive it", () => {
    const { batchUpdateCalls } = stubSheetWithDataRows(2);

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).row(5).delete();
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        deleteDimension: {
          range: {
            sheetId: 111,
            dimension: "ROWS",
            startIndex: 5,
            endIndex: 6,
          },
        },
      },
    ]);
  });

  it("counts a row appended and then deleted as neither, since the two cancel before the flush", () => {
    stubSheetWithDataRows(1);

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).appendDataRow().delete();

    expect(() => raw.sheet(111).topRow.delete()).toThrowError(/last data row/);
  });

  it("does not change another sheet's data-row count when a row delete is queued", () => {
    stubSheetsService({
      sheets: [
        { sheetId: 111, title: "Records", table: { endRowIndex: 11 } },
        { sheetId: 222, title: "Entries", table: { endRowIndex: 6 } },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    const unitsBefore = raw.sheet(222).dataRowCountAfterFlush;

    raw.sheet(111).row(5).delete();

    expect(raw.sheet(222).dataRowCountAfterFlush).toBe(unitsBefore);
    expect(raw.sheet(111).dataRowCountAfterFlush).toBe(6);
  });
});

describe("RowRaw.rowIsActive", () => {
  it("makes an appended row active and grows the table end before the flush", () => {
    stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    const endBefore = raw.sheet(111).activeTable.endRowIndex;

    const row = raw.sheet(111).appendDataRow();

    expect(row.rowIsActive()).toBe(true);
    expect(raw.sheet(111).activeTable.endRowIndex).toBe(endBefore + 1);
  });

  it("drops a removed row from the working view before the flush, leaving table indexes unmoved", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({ 4: ["kept"], 5: ["deleted"] }),
          table: { endRowIndex: 11 },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).row(5).gatherFetchFull();
    raw.fetchAllGathered();
    expect(raw.sheet(111).row(5).rowIsActive()).toBe(true);
    const endBefore = raw.sheet(111).activeTable.endRowIndex;

    raw.sheet(111).row(5).delete();

    expect(raw.sheet(111).row(5).rowIsActive()).toBe(false);
    expect(raw.sheet(111).activeTable.endRowIndex).toBe(endBefore);
    expect(raw.sheet(111).dataRowCountAfterFlush).toBe(6);
  });

  it("keeps pre-flush row indexes after a flushed delete, so the removed row stays inactive at its old index", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({ 4: ["kept"], 5: ["deleted"] }),
          table: { endRowIndex: 11 },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).row(5).gatherFetchFull();
    raw.fetchAllGathered();
    raw.sheet(111).row(5).delete();
    raw.batchUpdateGSheets();

    expect(raw.sheet(111).row(5).rowIsActive()).toBe(false);
    expect(raw.sheet(111).rowIndexesAreStale).toBe(true);
    expect(() => raw.sheet(111).activeTable.endRowIndex).toThrow(
      /Row indexes are stale/,
    );
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
    raw.sheet(111).topRow.gatherFetchFull();
    raw
      .sheet(111)
      .row(topDataRowIndex + 1)
      .gatherFetchFull();
    raw.fetchAllGathered();
    return raw;
  }

  it("leaves a row queued for delete inactive after a re-fetch that returns it", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.sheet(111).topRow.delete();
    raw.sheet(111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.sheet(111).topRow.rowIsActive()).toBe(false);
  });

  it("leaves the same row inactive when the re-fetch was a full row, so finalize backfilled", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.sheet(111).topRow.delete();
    raw.sheet(111).topRow.gatherFetchFull();
    expect(() => raw.fetchAllGathered()).not.toThrow();
    expect(raw.sheet(111).topRow.rowIsActive()).toBe(false);
  });

  it("leaves the same row inactive after a full-column fetch that covers it", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.sheet(111).topRow.delete();
    raw.sheet(111).column(1).gatherFetchFull();
    expect(() => raw.fetchAllGathered()).not.toThrow();
    expect(raw.sheet(111).topRow.rowIsActive()).toBe(false);
  });

  it("still supplies Table column facts from a top data row queued for delete", () => {
    stubTwoDataRows([
      "r:lse:1",
      { value: 42, isFormula: true, numberFormatType: "CURRENCY" },
    ]);

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).topRow.gatherFetchFull();
    raw.fetchAllGathered(true);
    raw.sheet(111).topRow.delete();
    raw.sheet(111).topRow.gatherFetchFull();
    expect(() => raw.fetchAllGathered(true)).not.toThrow();

    const column = raw.sheetMeta(111).column(1);
    expect(column.activeIsFormula).toBe(true);
    expect(column.activeNumberFormatType).toBe("CURRENCY");
    expect(column.activeTopValue).toBe(42);
  });

  it("still answers whether the top data row is blank from those facts", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.sheet(111).topRow.delete();
    raw.sheet(111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.sheet(111).topDataRowIsBlank()).toBe(false);
  });

  it("keeps a queued value update after a re-fetch of that cell", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.sheet(111).topRow.cell(1).updateValue("queued");
    raw.sheet(111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.sheet(111).topRow.valueOrEmpty(1)).toBe("queued");
  });

  it("keeps a queued value fill after a re-fetch of a covered cell", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.sheet(111).column(1).updateAllCells({ value: "filled" });
    raw.sheet(111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.sheet(111).topRow.valueOrEmpty(1)).toBe("filled");
  });

  it("lets the most recently queued value fill win when several cover one cell", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.sheet(111).column(1).updateAllCells({ value: "first" });
    raw.sheet(111).column(1).updateAllCells({ value: "second" });
    raw.sheet(111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.sheet(111).topRow.valueOrEmpty(1)).toBe("second");
  });

  it("lets the cell's own queued value win over a covering fill", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.sheet(111).column(1).updateAllCells({ value: "filled" });
    raw.sheet(111).topRow.cell(1).updateValue("queued");
    raw.sheet(111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.sheet(111).topRow.valueOrEmpty(1)).toBe("queued");
  });

  it("takes the live value when the cell has no queued value", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.sheet(111).topRow.cell(1).updateValue("stale local");
    raw.discardQueuedChanges();
    raw.sheet(111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.sheet(111).topRow.valueOrEmpty(1)).toBe("live");
  });

  it("takes the live value when only a formula update is queued", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.sheet(111).topRow.cell(1).updateFormula("=A1");
    raw.sheet(111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.sheet(111).topRow.valueOrEmpty(1)).toBe("live");
  });

  it("takes the live value when only a formula fill is queued", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.sheet(111).column(1).updateAllFormulas("=A1");
    raw.sheet(111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.sheet(111).topRow.valueOrEmpty(1)).toBe("live");
  });

  it("integrates live values after the flush has cleared the queue", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.sheet(111).topRow.cell(1).updateValue("queued");
    raw.batchUpdateGSheets();
    raw.sheet(111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.sheet(111).topRow.valueOrEmpty(1)).toBe("live");
  });

  it("applies a value queued before the row was fetched once that row is fetched", () => {
    stubTwoDataRows();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).topRow.cell(1).updateValue("queued");
    raw.sheet(111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.sheet(111).topRow.valueOrEmpty(1)).toBe("queued");
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
    raw.sheet(111).updateTitle("Renamed");
    raw.fetchAllSheetProperties();

    expect(raw.sheet(111).title).toBe("Renamed");
  });

  it("keeps a queued Table name on the known Table and in the sheet's Tables after a re-fetch", () => {
    stubNamedTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).updateTableName("renamedRecords");
    raw.fetchAllSheetProperties();

    expect(raw.sheet(111).activeTable.name).toBe("renamedRecords");
    expect(raw.sheet(111).tables).toEqual([
      { tableId: "fake-table-111", name: "renamedRecords" },
    ]);
  });

  it("keeps a queued column type after a re-fetch of the sheet properties", () => {
    stubNamedTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheetMeta(111).column(1).updateColumnType("DOUBLE");
    raw.fetchAllSheetProperties();

    expect(raw.sheetMeta(111).column(1).activeColumnType).toBe("DOUBLE");
  });

  it("lets the last of two queued tab titles win after a re-fetch", () => {
    stubNamedTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).updateTitle("First");
    raw.sheet(111).updateTitle("Second");
    raw.fetchAllSheetProperties();

    expect(raw.sheet(111).title).toBe("Second");
  });

  it("lets the last of two queued Table names win after a re-fetch", () => {
    stubNamedTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).updateTableName("firstRecords");
    raw.sheet(111).updateTableName("secondRecords");
    raw.fetchAllSheetProperties();

    expect(raw.sheet(111).activeTable.name).toBe("secondRecords");
  });

  it("leaves another sheet's queued title alone on a re-fetch of one sheet", () => {
    stubNamedTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(222).updateTitle("Renamed");
    raw.fetchSheetUsedGrid(111);

    expect(raw.sheet(222).title).toBe("Renamed");
  });

  it("applies a sheet's queued title to that sheet only", () => {
    stubNamedTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).updateTitle("Renamed");
    raw.fetchAllSheetProperties();

    expect(raw.sheet(222).title).toBe("Entries");
  });

  it("integrates the live title and Table name after the flush has cleared the queue", () => {
    stubNamedTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).updateTitle("Renamed");
    raw.sheet(111).updateTableName("renamedRecords");
    raw.batchUpdateGSheets();
    raw.fetchAllSheetProperties();

    // The fake does not replay renames, so the live sheet still has the old ones.
    expect(raw.sheet(111).title).toBe("Records");
    expect(raw.sheet(111).activeTable.name).toBe("records");
  });
});

describe("CellRaw.updateValue", () => {
  it("sends a write to a row that was never fetched, since a write needs no fetched state", () => {
    const { batchUpdateCalls } = stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).row(5).cell(2).updateValue("Processing...");
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        updateCells: {
          range: {
            sheetId: 111,
            startRowIndex: 5,
            endRowIndex: 6,
            startColumnIndex: 2,
            endColumnIndex: 3,
          },
          rows: [
            {
              values: [{ userEnteredValue: { stringValue: "Processing..." } }],
            },
          ],
          fields: "userEnteredValue",
        },
      },
    ]);
  });

  it("leaves an unfetched row unreadable, so a forgotten fetch still fails loudly on read", () => {
    stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    const cell = raw.sheet(111).row(5).cell(2);
    cell.updateValue("Processing...");

    expect(() => cell.valueOrEmpty()).toThrowError(/does not have a value set/);
  });

  it("throws for a data row past the table's last row rather than writing off the grid", () => {
    stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();

    expect(() => raw.sheet(111).row(11).cell(2).updateValue("x")).toThrowError(
      /past the last row/,
    );
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
    raw.sheet(111).row(4).gatherFetchFull();
    raw.fetchAllGathered();
    const cell = raw.sheet(111).row(4).cell(1);
    cell.updateValue("new");

    expect(cell.valueOrEmpty()).toBe("new");
  });
});

describe("SheetRaw column insert", () => {
  function stubThreeColumnTable() {
    return stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({
            0: ["c:lse:aaa", "c:lse:bbb", "c:lse:ccc"],
            4: ["r:lse:1", "left", "right"],
          }),
          table: { endRowIndex: 11, endColumnIndex: 3 },
        },
      ],
    });
  }

  it("grows the exclusive end column for an insert at the Table end and does not mark columns stale", () => {
    stubThreeColumnTable();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    const insertedIndex = raw.sheetMeta(111).insertColumnAtEnd({
      idPrefix: "lse",
      header: "New",
    });
    raw.batchUpdateGSheets();

    expect(insertedIndex).toBe(3);
    expect(raw.sheet(111).activeTable.endColumnIndex).toBe(4);
    expect(raw.sheet(111).rowIndexesAreStale).toBe(false);
    expect(raw.sheet(111).activeTable.endRowIndex).toBe(11);
    expect(() =>
      raw.sheet(111).row(5).cell(2).updateValue("kept"),
    ).not.toThrow();
    expect(() =>
      raw.sheet(111).row(5).cell(insertedIndex).updateValue("new"),
    ).not.toThrow();
  });

  it("marks column indexes at and to the right of a mid-Table insert stale, and leaves indexes to the left writable", () => {
    stubThreeColumnTable();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).addSheetChangeToSave({
      action: "insertColumn",
      startColumnIndex: 1,
    });
    raw.batchUpdateGSheets();

    expect(raw.sheet(111).rowIndexesAreStale).toBe(false);
    expect(raw.sheet(111).activeTable.endRowIndex).toBe(11);
    expect(raw.sheet(111).activeTable.endColumnIndex).toBe(3);
    expect(() =>
      raw.sheet(111).row(5).cell(0).updateValue("left"),
    ).not.toThrow();
    expect(() => raw.sheet(111).row(5).cell(1).updateValue("mid")).toThrow(
      "Column index 1 is stale. First stale column index is 1.",
    );
    expect(() => raw.sheet(111).row(5).cell(2).updateValue("right")).toThrow(
      "Column index 2 is stale. First stale column index is 1.",
    );
  });

  function insertDimensionRequests(
    batchUpdateCalls: { requests?: object[] }[],
  ): object[] {
    return batchUpdateCalls
      .flatMap(({ requests = [] }) => requests)
      .filter((request) => "insertDimension" in request);
  }

  function insertColumnRequest(startIndex: number) {
    return {
      insertDimension: {
        range: {
          sheetId: 111,
          dimension: "COLUMNS",
          startIndex,
          endIndex: startIndex + 1,
        },
        inheritFromBefore: false,
      },
    };
  }

  it("sends two end inserts on one sheet as two insertDimension requests in queue order", () => {
    const { batchUpdateCalls } = stubThreeColumnTable();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    const first = raw
      .sheetMeta(111)
      .insertColumnAtEnd({ idPrefix: "lse", header: "First" });
    const second = raw
      .sheetMeta(111)
      .insertColumnAtEnd({ idPrefix: "lse", header: "Second" });
    raw.batchUpdateGSheets();

    expect([first, second]).toEqual([3, 4]);
    expect(raw.sheet(111).activeTable.endColumnIndex).toBe(5);
    expect(insertDimensionRequests(batchUpdateCalls)).toEqual([
      insertColumnRequest(3),
      insertColumnRequest(4),
    ]);
  });

  it("refuses a mid-Table insert queued beside another insert on that sheet, and sends only the first", () => {
    const { batchUpdateCalls } = stubThreeColumnTable();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheetMeta(111).insertColumnAtEnd({ idPrefix: "lse", header: "New" });

    expect(() =>
      raw.sheet(111).addSheetChangeToSave({
        action: "insertColumn",
        startColumnIndex: 1,
      }),
    ).toThrow(
      'Refusing to queue a column insert at 1 on "Records" (gid 111): it already has a column insert queued, so the next must land at the Table end, 4.',
    );
    raw.batchUpdateGSheets();

    expect(insertDimensionRequests(batchUpdateCalls)).toEqual([
      insertColumnRequest(3),
    ]);
  });

  it("refuses a second insert queued beside a mid-Table insert on that sheet, and sends only the first", () => {
    const { batchUpdateCalls } = stubThreeColumnTable();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).addSheetChangeToSave({
      action: "insertColumn",
      startColumnIndex: 1,
    });

    expect(() =>
      raw.sheetMeta(111).insertColumnAtEnd({ idPrefix: "lse", header: "New" }),
    ).toThrow(
      'Refusing to queue a column insert on "Records" (gid 111): a mid-Table column insert is already queued.',
    );
    raw.batchUpdateGSheets();

    expect(insertDimensionRequests(batchUpdateCalls)).toEqual([
      insertColumnRequest(1),
    ]);
  });

  it("refuses a second Table-end insert that skips past the end plus the inserts already queued", () => {
    stubThreeColumnTable();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).addSheetChangeToSave({
      action: "insertColumn",
      startColumnIndex: 3,
    });

    expect(() =>
      raw.sheet(111).addSheetChangeToSave({
        action: "insertColumn",
        startColumnIndex: 5,
      }),
    ).toThrow("the next must land at the Table end, 4.");
  });

  it("still reads a cell whose column index became stale, because reads do not consult the watermark", () => {
    stubThreeColumnTable();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).row(4).gatherFetchFull();
    raw.fetchAllGathered();
    raw.sheet(111).addSheetChangeToSave({
      action: "insertColumn",
      startColumnIndex: 1,
    });
    raw.batchUpdateGSheets();

    expect(raw.sheet(111).row(4).valueOrEmpty(1)).toBe("left");
  });
});

describe("SpreadsheetRaw.discardQueuedChanges", () => {
  it("sends nothing for changes queued before the discard", () => {
    const { batchUpdateCalls } = stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).row(5).delete();
    raw.discardQueuedChanges();
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls).toEqual([]);
    expect(raw.sheet(111).rowIndexesAreStale).toBe(false);
    expect(raw.sheet(111).activeTable.endRowIndex).toBe(11);
  });

  it("empties the spreadsheet and per-sheet write queues, so a later flush sends nothing", () => {
    const { batchUpdateCalls } = stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).row(5).delete();
    raw.sheet(111).requestSortGSheet({
      colIdxToSortBy: 0,
      sortOrder: "ASCENDING",
    });
    raw.gatherRawRequest(
      googleRawRequest({ updateTable: { table: { tableId: "t" } } }),
    );
    raw.discardQueuedChanges();
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls).toEqual([]);
  });

  it("still sends changes queued after the discard, so a failure handler can report status", () => {
    const { batchUpdateCalls } = stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).row(5).delete();
    raw.discardQueuedChanges();
    raw.sheet(111).appendDataRow();
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        appendCells: {
          sheetId: 111,
          tableId: "fake-table-111",
          rows: [{}],
          fields: "userEnteredValue",
        },
      },
    ]);
  });
});

describe("ColumnRaw.updateAllCells", () => {
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
            6: ["r:lse:3", "old"],
          }),
          table: { endRowIndex: 7 },
        },
      ],
    });
  }
  function fetchedColumn() {
    const raw = SpreadsheetRaw.init();
    raw.sheet(111).column(1).gatherFetchFull();
    raw.fetchAllGathered();
    return raw;
  }

  it("sends one repeatCell for the whole column instead of one write per row", () => {
    const { batchUpdateCalls } = stubFilledSheet();

    const raw = fetchedColumn();
    raw.sheet(111).column(1).updateAllCells({ value: "new" });
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        repeatCell: {
          range: {
            sheetId: 111,
            startRowIndex: 4,
            endRowIndex: 7,
            startColumnIndex: 1,
            endColumnIndex: 2,
          },
          cell: { userEnteredValue: { stringValue: "new" } },
          fields: "userEnteredValue",
        },
      },
    ]);
  });

  it("mirrors the fill into row state, so a read before the flush sees it", () => {
    stubFilledSheet();

    const raw = fetchedColumn();
    raw.sheet(111).column(1).updateAllCells({ value: "new" });

    expect(raw.sheet(111).column(1).valueArrOrEmpty).toEqual([
      "new",
      "new",
      "new",
    ]);
  });

  it("orders a per-cell write after the fill, so the cell wins", () => {
    const { batchUpdateCalls } = stubFilledSheet();

    const raw = fetchedColumn();
    raw.sheet(111).column(1).updateAllCells({ value: "filled" });
    raw.sheet(111).row(5).cell(1).updateValue("overridden");
    raw.batchUpdateGSheets();

    const requests = batchUpdateCalls[0]?.requests ?? [];
    expect(requests[0]?.repeatCell?.cell?.userEnteredValue).toEqual({
      stringValue: "filled",
    });
    expect(
      requests[1]?.updateCells?.rows?.[0]?.values?.[0]?.userEnteredValue,
    ).toEqual({
      stringValue: "overridden",
    });
  });

  it("carries a background colour alongside the value in the one fill request", () => {
    const { batchUpdateCalls } = stubFilledSheet();

    const raw = fetchedColumn();
    raw
      .sheet(111)
      .column(1)
      .updateAllCells({ value: "new", backgroundColor: lightGreen });
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        repeatCell: {
          range: {
            sheetId: 111,
            startRowIndex: 4,
            endRowIndex: 7,
            startColumnIndex: 1,
            endColumnIndex: 2,
          },
          cell: {
            userEnteredValue: { stringValue: "new" },
            userEnteredFormat: { backgroundColor: lightGreen },
          },
          fields: "userEnteredValue,userEnteredFormat.backgroundColor",
        },
      },
    ]);
  });

  it("leaves a row appended after the fill alone, since the fill's bound is snapshotted", () => {
    const { batchUpdateCalls } = stubFilledSheet();

    const raw = fetchedColumn();
    raw.sheet(111).column(1).updateAllCells({ value: "filled" });
    raw.sheet(111).appendDataRow();
    raw.batchUpdateGSheets();

    const fill = (batchUpdateCalls[0]?.requests ?? []).find(
      (r) => r.repeatCell,
    );
    expect(fill?.repeatCell?.range?.endRowIndex).toBe(7);
  });
});

describe("ColumnRaw.updateActiveCells", () => {
  function stubSelectionSheet() {
    return stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({
            0: ["c:lse:aaa", "c:lse:bbb"],
            4: ["r:lse:1", "old"],
            5: ["r:lse:2", "old"],
            6: ["r:lse:3", "old"],
            7: ["r:lse:4", "old"],
            8: ["r:lse:5", "old"],
          }),
          table: { endRowIndex: 9 },
        },
      ],
    });
  }
  function fetchedSelectionSheet() {
    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheetMeta(111).colIdRow.gatherFetchFull();
    raw.sheet(111).column(1).gatherFetchFull();
    raw.fetchAllGathered();
    return raw;
  }
  function fillRanges(
    calls: GoogleAppsScript.Sheets.Schema.BatchUpdateSpreadsheetRequest[],
  ) {
    return calls
      .flatMap((call) => call.requests ?? [])
      .filter((request) => request.repeatCell)
      .map((request) => ({
        startRowIndex: request.repeatCell?.range?.startRowIndex,
        endRowIndex: request.repeatCell?.range?.endRowIndex,
      }));
  }

  it("sends one request for a column whose active rows are all contiguous", () => {
    const { batchUpdateCalls } = stubSelectionSheet();

    const raw = fetchedSelectionSheet();
    raw.sheet(111).column(1).updateActiveCells({ value: "new" });
    raw.batchUpdateGSheets();

    expect(fillRanges(batchUpdateCalls)).toEqual([
      { startRowIndex: 4, endRowIndex: 9 },
    ]);
  });

  it("sends one request per contiguous run rather than one per row", () => {
    const { batchUpdateCalls } = stubSelectionSheet();

    const raw = fetchedSelectionSheet();
    raw.sheet(111).removeRowsExcept(4, 5, 8);
    raw.sheet(111).column(1).updateActiveCells({ value: "new" });
    raw.batchUpdateGSheets();

    expect(fillRanges(batchUpdateCalls)).toEqual([
      { startRowIndex: 4, endRowIndex: 6 },
      { startRowIndex: 8, endRowIndex: 9 },
    ]);
  });

  it("carries value and background colour together under a mask naming both", () => {
    const { batchUpdateCalls } = stubSelectionSheet();

    const raw = fetchedSelectionSheet();
    raw.sheet(111).removeRowsExcept(4);
    raw
      .sheet(111)
      .column(1)
      .updateActiveCells({ value: "new", backgroundColor: lightGreen });
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        repeatCell: {
          range: {
            sheetId: 111,
            startRowIndex: 4,
            endRowIndex: 5,
            startColumnIndex: 1,
            endColumnIndex: 2,
          },
          cell: {
            userEnteredValue: { stringValue: "new" },
            userEnteredFormat: { backgroundColor: lightGreen },
          },
          fields: "userEnteredValue,userEnteredFormat.backgroundColor",
        },
      },
    ]);
  });

  it("leaves values alone when only a background colour is written", () => {
    const { batchUpdateCalls } = stubSelectionSheet();

    const raw = fetchedSelectionSheet();
    raw.sheet(111).removeRowsExcept(4);
    raw.sheet(111).column(1).updateActiveCells({ backgroundColor: lightGreen });
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls[0]?.requests?.[0]?.repeatCell?.cell).toEqual({
      userEnteredFormat: { backgroundColor: lightGreen },
    });
    expect(raw.sheet(111).column(1).valueArrOrEmpty).toEqual(["old"]);
  });

  it("writes nothing when no row is active", () => {
    const { batchUpdateCalls } = stubSelectionSheet();

    const raw = fetchedSelectionSheet();
    raw.sheet(111).removeRowsExcept();
    raw.sheet(111).column(1).updateActiveCells({ value: "new" });
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls).toEqual([]);
  });

  it("mirrors the write into row state, so a read before the flush sees it", () => {
    stubSelectionSheet();

    const raw = fetchedSelectionSheet();
    raw.sheet(111).removeRowsExcept(4, 8);
    raw.sheet(111).column(1).updateActiveCells({ value: "new" });

    expect(raw.sheet(111).column(1).valueArrOrEmpty).toEqual(["new", "new"]);
  });
});

describe("SheetRaw.removeRowsExcept", () => {
  function stubPrunableSheet() {
    return stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({
            0: ["c:lse:aaa", "c:lse:bbb"],
            4: ["r:lse:1", "old"],
            5: ["r:lse:2", "old"],
            6: ["r:lse:3", "old"],
          }),
          table: { endRowIndex: 7 },
        },
      ],
    });
  }
  function fetchedPrunableSheet() {
    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheetMeta(111).colIdRow.gatherFetchFull();
    raw.sheet(111).column(1).gatherFetchFull();
    raw.fetchAllGathered();
    return raw;
  }

  it("drops every data row that was not kept", () => {
    stubPrunableSheet();

    const raw = fetchedPrunableSheet();
    raw.sheet(111).removeRowsExcept(5);

    expect(raw.sheet(111).rowIndexesActive).toEqual([5]);
  });

  it("keeps the uniform rows, so a column still resolves by its id afterwards", () => {
    stubPrunableSheet();

    const raw = fetchedPrunableSheet();
    raw.sheet(111).removeRowsExcept(5);

    expect(raw.sheetMeta(111).columnByActiveId("c:lse:bbb").colIndex).toBe(1);
  });

  it("makes a whole-column fill throw, so it can't overwrite the excluded rows", () => {
    stubPrunableSheet();

    const raw = fetchedPrunableSheet();
    raw.sheet(111).removeRowsExcept(5);

    expect(() =>
      raw.sheet(111).column(1).updateAllCells({ value: "new" }),
    ).toThrowError(/pruned to a selection/);
  });
});

describe("ColumnRaw.updateAllFormulas", () => {
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
            6: ["r:lse:3", "old"],
          }),
          table: { endRowIndex: 7 },
        },
      ],
    });
  }

  it("serializes pasteData PASTE_FORMULA once per fill and does not mirror into row state", () => {
    const { batchUpdateCalls } = stubFilledSheet();

    const raw = SpreadsheetRaw.init();
    raw.sheet(111).column(1).gatherFetchFull();
    raw.fetchAllGathered();
    raw.sheet(111).column(1).updateAllFormulas("=2+1");
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        pasteData: {
          coordinate: { sheetId: 111, rowIndex: 4, columnIndex: 1 },
          data: '"=2+1"\n"=2+1"\n"=2+1"',
          delimiter: "\t",
          type: "PASTE_FORMULA",
        },
      },
    ]);
    expect(raw.sheet(111).column(1).valueArrOrEmpty).toEqual([
      "old",
      "old",
      "old",
    ]);
  });

  it("quotes commas and quotes so a FILTER formula stays in one cell", () => {
    const { batchUpdateCalls } = stubFilledSheet();
    const formula = '=FILTER(A:A,A:A<>"")';

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).row(4).cell(1).updateFormula(formula);
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        pasteData: {
          coordinate: { sheetId: 111, rowIndex: 4, columnIndex: 1 },
          data: '"=FILTER(A:A,A:A<>"""")"',
          delimiter: "\t",
          type: "PASTE_FORMULA",
        },
      },
    ]);
  });

  it("keeps a pretty-printed formula in one quoted field", () => {
    const { batchUpdateCalls } = stubFilledSheet();
    const formula = "=2+SINGLE(\ntest[Number])";

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).row(4).cell(1).updateFormula(formula);
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls[0]?.requests?.[0]?.pasteData?.data).toBe(
      `"${formula}"`,
    );
  });

  it("sends pasteData then a colour repeatCell in the same batchUpdate", () => {
    const { batchUpdateCalls } = stubFilledSheet();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).column(1).updateAllFormulas("=2+1");
    raw.sheet(111).column(1).updateAllCells({ backgroundColor: lightGreen });
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        pasteData: {
          coordinate: { sheetId: 111, rowIndex: 4, columnIndex: 1 },
          data: '"=2+1"\n"=2+1"\n"=2+1"',
          delimiter: "\t",
          type: "PASTE_FORMULA",
        },
      },
      {
        repeatCell: {
          range: {
            sheetId: 111,
            startRowIndex: 4,
            endRowIndex: 7,
            startColumnIndex: 1,
            endColumnIndex: 2,
          },
          cell: { userEnteredFormat: { backgroundColor: lightGreen } },
          fields: "userEnteredFormat.backgroundColor",
        },
      },
    ]);
  });

  it("drops a previously queued value when a formula is written on the same cell", () => {
    const { batchUpdateCalls } = stubFilledSheet();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    const cell = raw.sheet(111).row(4).cell(1);
    cell.updateValue("new");
    cell.updateFormula("=2+1");
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        pasteData: {
          coordinate: { sheetId: 111, rowIndex: 4, columnIndex: 1 },
          data: '"=2+1"',
          delimiter: "\t",
          type: "PASTE_FORMULA",
        },
      },
    ]);
  });

  it("orders a per-cell formula paste after the fill paste, so the cell wins", () => {
    const { batchUpdateCalls } = stubFilledSheet();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).column(1).updateAllFormulas("=2+1");
    raw.sheet(111).row(5).cell(1).updateFormula("=9");
    raw.batchUpdateGSheets();

    const requests = batchUpdateCalls[0]?.requests ?? [];
    expect(requests[0]?.pasteData?.coordinate?.rowIndex).toBe(4);
    expect(requests[0]?.pasteData?.data).toBe('"=2+1"\n"=2+1"\n"=2+1"');
    expect(requests[1]?.pasteData).toEqual({
      coordinate: { sheetId: 111, rowIndex: 5, columnIndex: 1 },
      data: '"=9"',
      delimiter: "\t",
      type: "PASTE_FORMULA",
    });
  });
});

describe("CellRaw.updateBackgroundColor", () => {
  it("sends one updateCells request masking only the background colour, with no value", () => {
    const { batchUpdateCalls } = stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).row(5).cell(2).updateBackgroundColor(lightGreen);
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        updateCells: {
          range: {
            sheetId: 111,
            startRowIndex: 5,
            endRowIndex: 6,
            startColumnIndex: 2,
            endColumnIndex: 3,
          },
          rows: [
            {
              values: [{ userEnteredFormat: { backgroundColor: lightGreen } }],
            },
          ],
          fields: "userEnteredFormat.backgroundColor",
        },
      },
    ]);
  });

  it("collapses a value and a colour on one cell into a single request masking both", () => {
    const { batchUpdateCalls } = stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).row(5).cell(2).updateValue("2026-09-05 10:00:00");
    raw.sheet(111).row(5).cell(2).updateBackgroundColor(lightGreen);
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        updateCells: {
          range: {
            sheetId: 111,
            startRowIndex: 5,
            endRowIndex: 6,
            startColumnIndex: 2,
            endColumnIndex: 3,
          },
          rows: [
            {
              values: [
                {
                  userEnteredValue: { stringValue: "2026-09-05 10:00:00" },
                  userEnteredFormat: { backgroundColor: lightGreen },
                },
              ],
            },
          ],
          fields: "userEnteredValue,userEnteredFormat.backgroundColor",
        },
      },
    ]);
  });

  it("leaves a value queued for the cell intact when the colour is queued after it", () => {
    const { batchUpdateCalls } = stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).row(5).cell(2).updateValue("kept");
    raw.sheet(111).row(5).cell(2).updateBackgroundColor(lightGreen);
    raw.batchUpdateGSheets();

    const values =
      batchUpdateCalls[0]?.requests?.[0]?.updateCells?.rows?.[0]?.values;
    expect(values?.[0]?.userEnteredValue).toEqual({ stringValue: "kept" });
  });

  it("leaves the cell unreadable, since the read path never fetches colour", () => {
    stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    const cell = raw.sheet(111).row(5).cell(2);
    cell.updateBackgroundColor(lightGreen);

    expect(cell.isActive).toBe(false);
    expect(() => cell.valueOrEmpty()).toThrowError(/does not have a value set/);
  });
});

describe("CellRaw.addCheckboxValidation", () => {
  it("sends one setDataValidation with a BOOLEAN condition over the cell", () => {
    const { batchUpdateCalls } = stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).row(5).cell(2).addCheckboxValidation();
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        setDataValidation: {
          range: {
            sheetId: 111,
            startRowIndex: 5,
            endRowIndex: 6,
            startColumnIndex: 2,
            endColumnIndex: 3,
          },
          rule: { condition: { type: "BOOLEAN" } },
        },
      },
    ]);
  });
});

// A mis-wired accessor still type-checks; the instance checks catch it.
describe("SpreadsheetRaw navigation", () => {
  it("gives each accessor the class its return type names", () => {
    const raw = SpreadsheetRaw.init();
    const sheet = raw.sheet(111);
    const sheetMeta = raw.sheetMeta(111);
    const column = sheet.column(0);
    const columnMeta = sheetMeta.column(0);

    assertType<IsExactly<typeof sheet, SheetRaw>>(true);
    assertType<IsExactly<typeof sheetMeta, SheetMetaRaw>>(true);
    assertType<IsExactly<typeof sheet.meta, SheetMetaRaw>>(true);
    assertType<IsExactly<typeof sheetMeta.primary, SheetRaw>>(true);
    assertType<IsExactly<typeof column, ColumnRaw>>(true);
    assertType<IsExactly<typeof columnMeta, ColumnMetaRaw>>(true);
    assertType<IsExactly<typeof column.sheet, SheetRaw>>(true);
    assertType<IsExactly<typeof columnMeta.sheet, SheetMetaRaw>>(true);
    assertType<IsExactly<typeof column.meta, ColumnMetaRaw>>(true);
    assertType<IsExactly<typeof columnMeta.primary, ColumnRaw>>(true);
    assertType<IsExactly<ReturnType<typeof sheet.row>, RowRaw>>(true);
    assertType<IsExactly<ReturnType<typeof sheet.rowCommon>, RowCommonRaw>>(
      true,
    );

    expect(sheet.meta).toBeInstanceOf(SheetMetaRaw);
    expect(sheetMeta.primary).toBeInstanceOf(SheetRaw);
    expect(column).toBeInstanceOf(ColumnRaw);
    expect(columnMeta).toBeInstanceOf(ColumnMetaRaw);
    expect(column.sheet).toBeInstanceOf(SheetRaw);
    expect(columnMeta.sheet).toBeInstanceOf(SheetMetaRaw);
    expect(column.meta).toBeInstanceOf(ColumnMetaRaw);
    expect(columnMeta.primary).toBeInstanceOf(ColumnRaw);
    expect(sheet.row(4)).toBeInstanceOf(RowRaw);
    expect(sheet.rowCommon(4)).toBeInstanceOf(RowRaw);
    expect(sheet.rowCommon(0)).toBeInstanceOf(UniformRowRaw);
  });
});

describe("Raw value types", () => {
  it("declares CellStateRaw as the cell value and nothing else", () => {
    assertType<IsExactly<CellStateRaw, { value: CellValue }>>(true);
  });

  it("declares the blank the wire can hold, with nothing validating it away", () => {
    assertType<
      IsExactly<ReturnType<CellRaw<"boolean">["valueOrEmpty"]>, boolean | "">
    >(true);
    assertType<IsExactly<ReturnType<RowRaw["valueOrEmpty"]>, CellValue | "">>(
      true,
    );
    assertType<
      IsExactly<
        ReturnType<UniformRowRaw<"action">["valueOrEmpty"]>,
        boolean | ""
      >
    >(true);
    assertType<
      IsExactly<RowCellChange["backgroundColor"], RgbColor | undefined>
    >(true);
  });
});

describe("SpreadsheetRaw.findReplace", () => {
  function stubFilledSheet() {
    return stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({
            0: ["c:lse:aaa", "c:lse:bbb"],
            4: ["r:lse:1", "Currency"],
            5: ["r:lse:2", "Caretaking"],
            6: ["r:lse:3", "Currency"],
          }),
          table: { endRowIndex: 7 },
        },
      ],
    });
  }
  function fetchedColumn() {
    const raw = SpreadsheetRaw.init();
    raw.sheet(111).column(1).gatherFetchFull();
    raw.fetchAllGathered();
    return raw;
  }

  it("scopes a column's replace to that column's data rows", () => {
    const { batchUpdateCalls } = stubFilledSheet();

    const raw = fetchedColumn();
    raw.sheet(111).column(1).findReplace({
      find: "Currency",
      replacement: "Total",
      matchEntireCell: true,
    });
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        findReplace: {
          find: "Currency",
          replacement: "Total",
          matchEntireCell: true,
          range: {
            sheetId: 111,
            startRowIndex: 4,
            endRowIndex: 7,
            startColumnIndex: 1,
            endColumnIndex: 2,
          },
        },
      },
    ]);
  });

  it("scopes a sheet's replace by sheetId rather than by range", () => {
    const { batchUpdateCalls } = stubFilledSheet();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).findReplace({ find: "Currency", replacement: "Total" });
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        findReplace: {
          find: "Currency",
          replacement: "Total",
          sheetId: 111,
        },
      },
    ]);
  });

  it("carries an allSheets scope straight through", () => {
    const { batchUpdateCalls } = stubFilledSheet();

    const raw = SpreadsheetRaw.init();
    raw.findReplace({
      find: "Currency",
      replacement: "Total",
      scope: { allSheets: true },
      includeFormulas: true,
    });
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        findReplace: {
          find: "Currency",
          replacement: "Total",
          includeFormulas: true,
          allSheets: true,
        },
      },
    ]);
  });

  it("sends after the per-cell writes, whose text it would otherwise miss", () => {
    const { batchUpdateCalls } = stubFilledSheet();

    const raw = fetchedColumn();
    raw.findReplace({
      find: "Currency",
      replacement: "Total",
      scope: { allSheets: true },
    });
    raw.sheet(111).row(5).cell(1).updateValue("Currency");
    raw.sheet(111).row(6).delete();
    raw.batchUpdateGSheets();

    const requests = batchUpdateCalls[0]?.requests ?? [];
    expect(requests.map((request) => Object.keys(request)[0])).toEqual([
      "updateCells",
      "findReplace",
      "deleteDimension",
    ]);
  });

  it("leaves fetched values readable until the flush actually sends", () => {
    stubFilledSheet();

    const raw = fetchedColumn();
    raw
      .sheet(111)
      .column(1)
      .findReplace({ find: "Currency", replacement: "Total" });

    expect(raw.sheet(111).column(1).valueArrOrEmpty).toEqual([
      "Currency",
      "Caretaking",
      "Currency",
    ]);
  });

  it("makes a read after the flush throw rather than return a pre-replace value", () => {
    stubFilledSheet();

    const raw = fetchedColumn();
    raw
      .sheet(111)
      .column(1)
      .findReplace({ find: "Currency", replacement: "Total" });
    raw.batchUpdateGSheets();

    expect(() => raw.sheet(111).row(4).cell(1).valueOrEmpty()).toThrowError(
      /went stale when a findReplace was sent/,
    );
  });

  it("leaves fetched values alone when no findReplace was queued", () => {
    stubFilledSheet();

    const raw = fetchedColumn();
    raw.sheet(111).row(4).cell(1).updateValue("Total");
    raw.batchUpdateGSheets();

    expect(raw.sheet(111).column(1).valueArrOrEmpty).toEqual([
      "Total",
      "Caretaking",
      "Currency",
    ]);
  });

  it("makes the values readable again after a re-fetch", () => {
    stubFilledSheet();

    const raw = fetchedColumn();
    raw
      .sheet(111)
      .column(1)
      .findReplace({ find: "Currency", replacement: "Total" });
    raw.batchUpdateGSheets();
    raw.sheet(111).column(1).gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.sheet(111).column(1).valueArrOrEmpty).toEqual([
      "Currency",
      "Caretaking",
      "Currency",
    ]);
  });

  it("discards a queued replace alongside every other change", () => {
    const { batchUpdateCalls } = stubFilledSheet();

    const raw = SpreadsheetRaw.init();
    raw.findReplace({
      find: "Currency",
      replacement: "Total",
      scope: { allSheets: true },
    });
    raw.discardQueuedChanges();
    raw.batchUpdateGSheets();

    expect(batchUpdateCalls).toEqual([]);
  });
});

describe("SheetMetaRaw.activeColumnIds", () => {
  function fetchedColumnIdSheet(columnIdRow: FakeCell[]) {
    stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({
            0: columnIdRow,
            4: [],
          }),
          table: { endRowIndex: topDataRowIndex + 1, endColumnIndex: 2 },
        },
      ],
    });
    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheetMeta(111).colIdRow.gatherFetchFull();
    raw.fetchAllGathered();
    return raw.sheetMeta(111);
  }

  it("throws when a Table column-ID cell is a number, boolean, or date", () => {
    expect(
      () => fetchedColumnIdSheet(["c:lse:aaa", 42]).activeColumnIds,
    ).toThrow(/Records.*column index 1/);
    expect(
      () => fetchedColumnIdSheet(["c:lse:aaa", true]).activeColumnIds,
    ).toThrow(/Records.*column index 1/);
    expect(
      () => fetchedColumnIdSheet(["c:lse:aaa", 44927]).activeColumnIds,
    ).toThrow(/Records.*column index 1/);
  });

  it("treats a blank Table column-ID cell as missing rather than a type error", () => {
    expect(fetchedColumnIdSheet(["c:lse:aaa", ""]).activeColumnIds).toEqual([
      "c:lse:aaa",
    ]);
  });

  it("ignores a non-string past the Table", () => {
    expect(
      fetchedColumnIdSheet(["c:lse:aaa", "c:lse:bbb", 42]).activeColumnIds,
    ).toEqual(["c:lse:aaa", "c:lse:bbb"]);
  });

  it("throws from lookup by ID when a sibling Table cell is not text", () => {
    const sheet = fetchedColumnIdSheet(["c:lse:aaa", false]);
    expect(() => sheet.columnByActiveId("c:lse:aaa")).toThrow(
      /Records.*column index 1/,
    );
  });

  it("throws from fill-missing when a Table column-ID cell is not text", () => {
    expect(() =>
      fetchedColumnIdSheet(["", 42]).addMissingColumnIds("lse"),
    ).toThrow(/Records.*column index 1/);
  });
});

describe("SheetRaw.activeTable", () => {
  function fetchedSheet(endRowIndex: number) {
    stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          table: { endRowIndex },
        },
      ],
    });
    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    return raw.sheet(111);
  }

  it("throws when the exclusive end row is the first data row", () => {
    expect(() => fetchedSheet(topDataRowIndex).activeTable).toThrow(
      /Records.*at least one data row/,
    );
  });

  it("accepts a Table whose exclusive end is one past the first data row", () => {
    expect(fetchedSheet(topDataRowIndex + 1).activeTable.endRowIndex).toBe(
      topDataRowIndex + 1,
    );
  });
});

describe("ColumnMetaRaw.updateColumnType", () => {
  function stubTypedTable(
    table: Partial<NonNullable<FakeSheetProperties["table"]>> = {},
  ) {
    return stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({
            [tableHeaderRowIndex]: ["Name", "ID", "Amount"],
          }),
          table: {
            endRowIndex: tableEndRowIndex,
            endColumnIndex: startTableColIndex + 3,
            columnTypes: { [startTableColIndex]: "TEXT" },
            ...table,
          },
        },
      ],
    });
  }

  function fetchedRaw(): SpreadsheetRaw {
    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    return raw;
  }

  it("sends one full-list updateTable first in the batch, carrying every column's name and type", () => {
    const { batchUpdateCalls } = stubTypedTable();
    const raw = fetchedRaw();
    raw.sheet(111).appendDataRow();
    raw
      .sheet(111)
      .meta.column(startTableColIndex + 2)
      .updateColumnType("DOUBLE");
    raw
      .sheet(111)
      .meta.column(startTableColIndex + 1)
      .updateColumnType("TEXT");
    raw.batchUpdateGSheets();

    const requests = batchUpdateCalls[0]?.requests ?? [];
    expect(requests.filter((request) => request.updateTable)).toHaveLength(1);
    expect(requests[0]).toEqual({
      updateTable: {
        table: {
          tableId: "fake-table-111",
          columnProperties: [
            { columnIndex: 0, columnName: "Name", columnType: "TEXT" },
            { columnIndex: 1, columnName: "ID", columnType: "TEXT" },
            { columnIndex: 2, columnName: "Amount", columnType: "DOUBLE" },
          ],
        },
        fields: "columnProperties",
      },
    });
    expect(requests[1]?.appendCells).toBeDefined();
  });

  it("keeps an untouched column's name and type through the full-list replace", () => {
    stubTypedTable();
    const raw = fetchedRaw();
    raw
      .sheet(111)
      .meta.column(startTableColIndex + 2)
      .updateColumnType("DOUBLE");
    raw.batchUpdateGSheets();
    raw.fetchAllSheetProperties();

    expect(
      raw.sheet(111).meta.column(startTableColIndex).activeColumnType,
    ).toBe("TEXT");
    expect(
      raw.sheet(111).meta.column(startTableColIndex + 1).activeColumnType,
    ).toBeUndefined();
    expect(
      raw.sheet(111).meta.column(startTableColIndex + 2).activeColumnType,
    ).toBe("DOUBLE");
  });

  it("sends a sibling column's type back unchanged when it is one the framework does not name", () => {
    const { batchUpdateCalls } = stubTypedTable({
      columnTypes: { [startTableColIndex]: "FUTURE_CHIP" },
    });
    const raw = fetchedRaw();
    raw
      .sheet(111)
      .meta.column(startTableColIndex + 2)
      .updateColumnType("DOUBLE");
    raw.batchUpdateGSheets();

    expect(
      batchUpdateCalls[0]?.requests?.[0]?.updateTable?.table?.columnProperties,
    ).toEqual([
      { columnIndex: 0, columnName: "Name", columnType: "FUTURE_CHIP" },
      { columnIndex: 1, columnName: "ID" },
      { columnIndex: 2, columnName: "Amount", columnType: "DOUBLE" },
    ]);
  });

  it("fake: a columnProperties update replaces the whole list, so a column left out loses its type", () => {
    stubTypedTable();
    const raw = fetchedRaw();
    raw.gatherRawRequest(
      googleRawRequest({
        updateTable: {
          table: {
            tableId: "fake-table-111",
            columnProperties: [
              { columnIndex: 2, columnName: "Amount", columnType: "DOUBLE" },
            ],
          },
          fields: "columnProperties",
        },
      }),
    );
    raw.batchUpdateGSheets();
    raw.fetchAllSheetProperties();

    expect(
      raw.sheet(111).meta.column(startTableColIndex).activeColumnType,
    ).toBeUndefined();
    expect(
      raw.sheet(111).meta.column(startTableColIndex + 2).activeColumnType,
    ).toBe("DOUBLE");
  });

  it("fake: rejects a columnProperties update carrying a column with no columnName", () => {
    stubTypedTable();
    const raw = fetchedRaw();
    raw.gatherRawRequest(
      googleRawRequest({
        updateTable: {
          table: {
            tableId: "fake-table-111",
            columnProperties: [{ columnIndex: 2, columnType: "DOUBLE" }],
          },
          fields: "columnProperties",
        },
      }),
    );

    expect(() => raw.batchUpdateGSheets()).toThrow(/columnName/);
  });

  it("refuses before sending anything when a column on the Table has a validation rule, naming the Table and those columns", () => {
    const { batchUpdateCalls } = stubTypedTable({
      columnValidationValues: { [startTableColIndex + 1]: ["a", "b"] },
      columnValidationConditionTypes: {
        [startTableColIndex + 1]: "ONE_OF_LIST",
      },
    });
    const raw = fetchedRaw();
    raw
      .sheet(111)
      .meta.column(startTableColIndex + 2)
      .updateColumnType("DOUBLE");

    expect(() => raw.batchUpdateGSheets()).toThrow(
      /fake-table-111.*Records.*ID/,
    );
    expect(batchUpdateCalls).toEqual([]);
  });

  it("refuses when the same flush inserts a column on that sheet", () => {
    const { batchUpdateCalls } = stubTypedTable();
    const raw = fetchedRaw();
    raw
      .sheet(111)
      .meta.column(startTableColIndex + 2)
      .updateColumnType("DOUBLE");
    raw.sheet(111).addSheetChangeToSave({
      action: "insertColumn",
      startColumnIndex: startTableColIndex + 3,
    });

    expect(() => raw.batchUpdateGSheets()).toThrow(/inserts a column/);
    expect(batchUpdateCalls).toEqual([]);
  });

  it("refuses a second update after a flush until the Table is refetched", () => {
    const { batchUpdateCalls } = stubTypedTable();
    const raw = fetchedRaw();
    raw
      .sheet(111)
      .meta.column(startTableColIndex + 2)
      .updateColumnType("DOUBLE");
    raw.batchUpdateGSheets();
    raw
      .sheet(111)
      .meta.column(startTableColIndex + 1)
      .updateColumnType("TEXT");

    expect(() => raw.batchUpdateGSheets()).toThrow(
      /no fetched column properties/,
    );
    expect(batchUpdateCalls).toHaveLength(1);
  });
});
