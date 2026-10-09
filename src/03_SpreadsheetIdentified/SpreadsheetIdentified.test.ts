import { beforeEach, describe, expect, it } from "vitest";

import { getColumnTraitByName } from "../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import { getTableTraitByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import type {
  Value,
  VnToCvn,
} from "../01_SpreadsheetSchema/configReaders/valueSchemas";
import {
  colIdRowIndex,
  extraTablesSheet,
  itemTableId,
  layoutGid,
  layoutSheet,
  layoutTableId,
  logGid,
  logTableId,
  managedLayoutTableNames,
  misplacedTableSheet,
  ownColumnId,
  placedTableSheet,
  recordedGridRanges,
  tableEndRowIndex,
  tableHeaderRowIndex,
  thrownMessage,
  topDataRowIndex,
} from "../02_SpreadsheetRaw/spreadsheetRawTestSupport";
import {
  buildGridRows,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import {
  blankTableConfigRow,
  filledTableConfigRow,
  stubTableConfigSheet,
  tableConfigColumnIdRow,
  tableConfigGid,
} from "../testSupport/fakeTableConfigSheet";
import { assertType, type IsExactly } from "../testSupport/typeAssertions";
import { CellIdentified } from "./CellIdentified";
import { SpreadsheetBaseIdentified } from "./ClassBases/SpreadsheetBaseIdentified";
import type { FetchTargetIdentified } from "./ClassTypes/StateIdentified";
import { ColumnIdentified } from "./ColumnIdentified";
import { HeadRowIdentified } from "./HeadRowIdentified";
import { RowIdentified } from "./RowIdentified";
import { SpreadsheetIdentified } from "./SpreadsheetIdentified";
import { TableIdentified } from "./TableIdentified";

const itemGid = getTableTraitByName("item", "sheetGid");
const itemIdColumnId = getColumnTraitByName("item", "id", "columnId");

// A mis-wired accessor still type-checks; the instance checks catch it.
describe("SpreadsheetIdentified navigation", () => {
  function stubItemTable(): void {
    stubSheetsService({
      sheets: [placedTableSheet({ sheetId: itemGid, title: "Item" })],
    });
  }

  it("gives each accessor the class its return type names", () => {
    stubItemTable();
    const ssi = new SpreadsheetIdentified(
      SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
    );
    ssi.raw.fetchAllSheetProperties();
    const table = ssi.table(itemTableId);
    const tableOnSheet = ssi.tableOnSheet(itemGid);
    const column = table.column(itemIdColumnId);
    const headRow = table.headRow("action");
    const headCell = column.headCell("action");

    assertType<IsExactly<typeof table, TableIdentified>>(true);
    assertType<IsExactly<typeof tableOnSheet, TableIdentified>>(true);
    assertType<IsExactly<typeof column, ColumnIdentified>>(true);
    assertType<IsExactly<typeof column.table, TableIdentified>>(true);
    assertType<IsExactly<ReturnType<typeof table.row>, RowIdentified>>(true);
    assertType<IsExactly<typeof headRow, HeadRowIdentified<"action">>>(true);
    assertType<
      IsExactly<ReturnType<typeof table.headRowByIndex>, HeadRowIdentified>
    >(true);
    assertType<
      IsExactly<typeof headCell, CellIdentified<"boolean" | "string">>
    >(true);
    assertType<
      IsExactly<ReturnType<typeof headCell.valueOrEmpty>, boolean | string>
    >(true);

    expect(table).toBeInstanceOf(TableIdentified);
    expect(tableOnSheet).toBeInstanceOf(TableIdentified);
    expect(column).toBeInstanceOf(ColumnIdentified);
    expect(column.table).toBeInstanceOf(TableIdentified);
    expect(table.row(0)).toBeInstanceOf(RowIdentified);
    expect(headRow).toBeInstanceOf(HeadRowIdentified);
    expect(table.headRowByIndex(-2)).toBeInstanceOf(HeadRowIdentified);
    expect(headCell).toBeInstanceOf(CellIdentified);
  });

  it("offers no profile: descriptive facts stay at Raw", () => {
    assertType<
      IsExactly<
        Extract<"profile", keyof TableIdentified | keyof ColumnIdentified>,
        never
      >
    >(true);
  });

  it("offers no Meta view", () => {
    assertType<
      IsExactly<
        Extract<
          "meta" | "sheetMeta",
          | keyof SpreadsheetIdentified
          | keyof TableIdentified
          | keyof ColumnIdentified
        >,
        never
      >
    >(true);
  });

  it("reaches the Table back through each row's and column's table getter", () => {
    stubItemTable();
    const ssi = new SpreadsheetIdentified(
      SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
    );
    ssi.raw.fetchAllSheetProperties();
    const table = ssi.table(itemTableId);
    const row = table.row(0);
    const column = table.column(itemIdColumnId);
    const headRow = table.headRow("header");

    assertType<IsExactly<typeof row.table, TableIdentified>>(true);
    assertType<IsExactly<typeof column.table, TableIdentified>>(true);
    assertType<IsExactly<typeof headRow.table, TableIdentified>>(true);

    expect(row.table).toBeInstanceOf(TableIdentified);
    expect(column.table).toBeInstanceOf(TableIdentified);
    expect(row.table).toEqual(table);
    expect(column.table).toEqual(table);
    expect(headRow.table).toBeInstanceOf(TableIdentified);
    expect(headRow.table).toEqual(table);
  });
});

const valueTypesGid = getTableTraitByName("valueTypes", "sheetGid");
const valueTypesIdColumnId = getColumnTraitByName(
  "valueTypes",
  "id",
  "columnId",
);
const checkboxColumnId = getColumnTraitByName(
  "valueTypes",
  "checkbox",
  "columnId",
);
const filledRowIndex = 0;
const blankRowIndex = 1;

// Body row 1 (sheet row 5) is the blank row; its checkbox is untouched, so it reads blank not false.
function stubValueTypesWithBlankRow() {
  return stubSheetsService({
    sheets: [
      {
        sheetId: valueTypesGid,
        title: "Value Types",
        rows: buildGridRows({
          0: [valueTypesIdColumnId, checkboxColumnId],
          3: ["ID", "Checkbox"],
          4: ["r:vty:row4", true],
          5: [null, null],
        }),
        table: { endRowIndex: 6 },
      },
    ],
  });
}

function fetchedValueTypesSheet(): TableIdentified {
  const ssi = new SpreadsheetIdentified(
    SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
  );
  const sheet = ssi.tableOnSheet(valueTypesGid);
  sheet.column(valueTypesIdColumnId).prepFetchFull();
  sheet.column(checkboxColumnId).prepFetchFull();
  ssi.fetchAllPrepped();
  return sheet;
}

describe("Identified value accessors", () => {
  beforeEach(() => {
    stubValueTypesWithBlankRow();
  });

  it("throws from CellIdentified.valueNotEmpty on a blank cell, naming the column id and the row", () => {
    const cell = fetchedValueTypesSheet()
      .column(valueTypesIdColumnId)
      .cell(blankRowIndex);

    expect(() => cell.valueNotEmpty()).toThrowError(
      new RegExp(`${valueTypesIdColumnId}.*row 6`),
    );
  });

  it("returns the empty string from CellIdentified.valueOrEmpty on that same cell", () => {
    const cell = fetchedValueTypesSheet()
      .column(valueTypesIdColumnId)
      .cell(blankRowIndex);

    expect(cell.valueOrEmpty()).toBe("");
  });

  it("reads a specifically fetched cell that Sheets omitted as empty, not unfetched", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: valueTypesGid,
          title: "Value Types",
          rows: buildGridRows({
            0: [valueTypesIdColumnId, checkboxColumnId],
            3: ["ID", "Checkbox"],
            4: ["r:vty:row4", true],
            5: [null, null],
          }),
          rowsWithNoGridData: [5],
          table: { endRowIndex: 6 },
        },
      ],
    });
    const ssi = new SpreadsheetIdentified(
      SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
    );
    const sheet = ssi.tableOnSheet(valueTypesGid);
    sheet.column(valueTypesIdColumnId).prepFetchSpecific([blankRowIndex]);
    ssi.fetchAllPrepped();

    expect(
      sheet.column(valueTypesIdColumnId).cell(blankRowIndex).valueOrEmpty(),
    ).toBe("");
  });

  it("reads a filled cell identically through both forms", () => {
    const cell = fetchedValueTypesSheet()
      .column(valueTypesIdColumnId)
      .cell(filledRowIndex);

    expect(cell.valueNotEmpty()).toBe("r:vty:row4");
    expect(cell.valueOrEmpty()).toBe("r:vty:row4");
  });

  it("reads an untouched checkbox as unchecked through every accessor", () => {
    const sheet = fetchedValueTypesSheet();
    const column = new ColumnIdentified<"checkbox">({
      ...sheet.tableIdentifiedProps,
      columnId: checkboxColumnId,
    });

    expect(column.valueOrEmpty(blankRowIndex)).toBe(false);
    expect(column.valueNotEmpty(blankRowIndex)).toBe(false);
    expect(column.valueNotEmpty(filledRowIndex)).toBe(true);
    expect(column.valueArrOrEmpty).toEqual([true, false]);
    expect(column.valueArrNotEmpty).toEqual([true, false]);
    assertType<IsExactly<ReturnType<typeof column.valueNotEmpty>, boolean>>(
      true,
    );
    assertType<IsExactly<ReturnType<typeof column.valueOrEmpty>, boolean>>(
      true,
    );
  });

  it("throws from ColumnIdentified.valueNotEmpty and returns empty from valueOrEmpty", () => {
    const column = fetchedValueTypesSheet().column(valueTypesIdColumnId);

    expect(() => column.valueNotEmpty(blankRowIndex)).toThrowError(/is empty/);
    expect(column.valueOrEmpty(blankRowIndex)).toBe("");
  });

  it("throws from RowIdentified.valueNotEmpty and returns empty from RowIdentified.valueOrEmpty", () => {
    const row = fetchedValueTypesSheet().row(blankRowIndex);

    expect(() => row.valueNotEmpty(valueTypesIdColumnId)).toThrowError(
      /is empty/,
    );
    expect(row.valueOrEmpty(valueTypesIdColumnId)).toBe("");
  });

  it("throws from valueArrNotEmpty when a fetched cell is blank, but not from the blank-tolerant forms", () => {
    const column = fetchedValueTypesSheet().column(valueTypesIdColumnId);

    expect(() => column.valueArrNotEmpty).toThrowError(/is empty/);
    expect(column.valueArrOrEmpty).toEqual(["r:vty:row4", ""]);
    expect(column.valueArrFilterEmpty).toEqual(["r:vty:row4"]);
  });

  it("gives the checkbox value name a type with no blank in it", () => {
    assertType<IsExactly<Value<"checkbox">, boolean>>(true);
    assertType<IsExactly<Value<"boolean">, boolean | "">>(true);
  });

  // What toWireValue asserts rather than proves, proved here.
  it("sends the checkbox value name down to the boolean wire type", () => {
    assertType<IsExactly<VnToCvn<"checkbox">, "boolean">>(true);
    assertType<IsExactly<VnToCvn<"boolean">, "boolean">>(true);
    assertType<IsExactly<VnToCvn<"id">, "string">>(true);
  });
});

// Google omits a row nothing was ever written to, which is what "never read" looks like.
function stubTableConfigWithUnreadTopRow() {
  return stubTableConfigSheet({ 4: blankTableConfigRow }, [4]);
}

function fetchedTableConfig(): TableIdentified {
  const ssi = new SpreadsheetIdentified(
    SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
  );
  const sheet = ssi.tableOnSheet(tableConfigGid);
  sheet.topRow.prepFetchFull();
  ssi.fetchAllPrepped();
  return sheet;
}

function unfetchedTableConfig(): TableIdentified {
  const ssi = new SpreadsheetIdentified(
    SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
  );
  ssi.tableOnSheet(tableConfigGid).ensureColumnIdsAreFetched();
  return ssi.tableOnSheet(tableConfigGid);
}

describe("Identified head rows", () => {
  const actionSheetRow = 2;

  function fetchedHeadRows() {
    const { grid } = stubSheetsService({
      sheets: [
        {
          sheetId: valueTypesGid,
          title: "Value Types",
          rows: buildGridRows({
            0: [valueTypesIdColumnId, checkboxColumnId],
            [actionSheetRow]: ["Due", true],
            3: ["ID", "Checkbox"],
            4: ["r:vty:row4", true],
          }),
          table: { endRowIndex: 5 },
        },
      ],
    });
    const ssi = new SpreadsheetIdentified(
      SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
    );
    const table = ssi.tableOnSheet(valueTypesGid);
    table.headRow("action").prepFetchFull();
    table.headRow("header").prepFetchFull();
    table.headRow("groupHeading1").prepFetchFull();
    ssi.fetchAllPrepped();
    return { grid, ssi, table };
  }

  it("reads head cells by column ID, through the column and through the row", () => {
    const { table } = fetchedHeadRows();

    expect(
      table.column(checkboxColumnId).headCell("action").valueOrEmpty(),
    ).toBe(true);
    expect(
      table.headRow("groupHeading2").valueOrEmpty(valueTypesIdColumnId),
    ).toBe("Due");
    expect(table.headRow("header").valueOrEmpty(checkboxColumnId)).toBe(
      "Checkbox",
    );
  });

  it("writes head cells by column ID, through the column and through the row", () => {
    const { grid, ssi, table } = fetchedHeadRows();

    table.column(checkboxColumnId).headCell("action").updateValue(false);
    table.headRow("groupHeading2").updateValue(valueTypesIdColumnId, "Late");
    ssi.raw.batchUpdateGSheets();

    expect(grid.sheet(valueTypesGid).cell(actionSheetRow, 0)).toBe("Late");
    expect(grid.sheet(valueTypesGid).cell(actionSheetRow, 1)).toBe(false);
  });

  it("reads a checkbox column's blank head cell as blank text", () => {
    const { table } = fetchedHeadRows();

    expect(
      table.column(checkboxColumnId).headCell("groupHeading1").valueOrEmpty(),
    ).toBe("");
  });

  it("finds the row at an index with every role it holds", () => {
    const { table } = fetchedHeadRows();

    expect(table.headRowByIndex(-2).roles).toEqual(["action", "groupHeading2"]);
    expect(() => table.headRowByIndex(0)).toThrow(/not a head row/);
  });
});

describe("TableIdentified.addMissingColumnIds", () => {
  it("writes a column ID with the Table's prefix into each blank column-ID cell", () => {
    const { grid } = stubSheetsService({
      sheets: [
        {
          sheetId: valueTypesGid,
          title: "Value Types",
          rows: buildGridRows({
            0: [valueTypesIdColumnId, ""],
            3: ["ID", "Checkbox"],
            4: ["r:vty:row4", true],
          }),
          table: { endRowIndex: 5 },
        },
      ],
    });
    const ssi = new SpreadsheetIdentified(
      SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
    );
    const table = ssi.tableOnSheet(valueTypesGid);

    const addedCount = table.ensureColumnIdsAreFetched().addMissingColumnIds();
    ssi.raw.batchUpdateGSheets();

    expect(addedCount).toBe(1);
    expect(grid.sheet(valueTypesGid).cell(0, 0)).toBe(valueTypesIdColumnId);
    expect(grid.sheet(valueTypesGid).cell(0, 1)).toMatch(
      new RegExp(`^c:${table.schema.idPrefix}:`),
    );
  });
});

describe("RowIdentified.isBlank / isReusable", () => {
  it("calls a row whose every non-formula cell is empty blank", () => {
    stubTableConfigSheet({ 4: blankTableConfigRow });

    expect(fetchedTableConfig().topRow.isBlank).toBe(true);
  });

  it("calls a row holding any non-formula value not blank", () => {
    stubTableConfigSheet({ 4: filledTableConfigRow });

    expect(fetchedTableConfig().topRow.isBlank).toBe(false);
  });

  it("calls a row nothing fetched not blank, since nothing read it", () => {
    stubTableConfigWithUnreadTopRow();

    expect(unfetchedTableConfig().topRow.isBlank).toBe(false);
  });

  it("makes a blank row reusable until an append reserves it", () => {
    stubTableConfigSheet({ 4: blankTableConfigRow });

    const sheet = fetchedTableConfig();
    expect(sheet.topRow.isReusable).toBe(true);

    sheet.appendRowDefault();
    expect(sheet.topRow.isReusable).toBe(false);
  });
});

describe("TableIdentified.hasNoData", () => {
  it("is true for a sheet whose one row is blank", () => {
    stubTableConfigSheet({ 4: blankTableConfigRow });

    expect(fetchedTableConfig().hasNoData).toBe(true);
  });

  it("is false while any row still holds data", () => {
    stubTableConfigSheet({ 4: blankTableConfigRow, 5: filledTableConfigRow });

    expect(fetchedTableConfig().hasNoData).toBe(false);
  });
});

const runItemGid = getTableTraitByName("runItem", "sheetGid");
const runItemColumnIds = (["id", "result", "runStatus"] as const).map(
  (columnName) => getColumnTraitByName("runItem", columnName, "columnId"),
);
const runStatusColumnId = getColumnTraitByName(
  "runItem",
  "runStatus",
  "columnId",
);

function runItemWithOneRow(topRow: (string | null)[]): TableIdentified {
  stubSheetsService({
    sheets: [
      {
        sheetId: runItemGid,
        title: "Run item",
        rows: buildGridRows({ 0: runItemColumnIds, 4: topRow }),
        table: { endRowIndex: 5 },
      },
    ],
  });
  const ssi = new SpreadsheetIdentified(
    SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(
      new Map([["runItem", new Set([runStatusColumnId])]]),
    ),
  );
  const sheet = ssi.tableOnSheet(runItemGid);
  sheet.topRow.prepFetchFull();
  ssi.fetchAllPrepped();
  return sheet;
}

describe("the blank test, on a sheet with a feedback column", () => {
  it("calls a row holding only a run status blank", () => {
    const sheet = runItemWithOneRow([null, null, "Succeeded"]);

    expect(sheet.topRow.isBlank).toBe(true);
    expect(sheet.hasNoData).toBe(true);
  });

  it("calls a row holding anything outside the feedback columns not blank", () => {
    const sheet = runItemWithOneRow([null, "a result", "Succeeded"]);

    expect(sheet.topRow.isBlank).toBe(false);
    expect(sheet.hasNoData).toBe(false);
  });

  it("leaves the feedback column out of a Table reached by its tableId", () => {
    const sheet = runItemWithOneRow([null, null, "Succeeded"]);
    const table = new SpreadsheetIdentified(
      sheet.spreadsheetIdentifiedProps,
    ).table(sheet.tableId);

    expect(table.blankTestColumnIds).not.toContain(runStatusColumnId);
    expect(table.topRow.isBlank).toBe(true);
  });
});

const computedGid = getTableTraitByName("computed", "sheetGid");
const amountColumnId = getColumnTraitByName("computed", "amount", "columnId");
const rowNumberColumnId = getColumnTraitByName(
  "computed",
  "rowNumber",
  "columnId",
);

const topRowRange = { startRowIndex: 4, endRowIndex: 5 };

describe("RowIdentified.clearValues", () => {
  it("empties every non-formula cell and touches no formula cell", () => {
    const { grid } = stubSheetsService({
      sheets: [
        {
          sheetId: computedGid,
          title: "Computed",
          rows: buildGridRows({
            0: [amountColumnId, rowNumberColumnId],
            4: [12, 5],
          }),
          table: { endRowIndex: 5 },
        },
      ],
    });

    const ssi = new SpreadsheetIdentified(
      SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
    );
    const sheet = ssi.tableOnSheet(computedGid);
    sheet.topRow.prepFetchFull();
    ssi.fetchAllPrepped();
    sheet.topRow.clearValues();
    ssi.raw.batchUpdateGSheets();

    expect(grid.sheet(computedGid).values(topRowRange)).toEqual([["", 5]]);
    expect(sheet.topRow.isBlank).toBe(true);
  });

  // The cell is cleared to a blank on the wire; the value name is what reads it back.
  it("leaves a cleared checkbox reading unchecked rather than blank", () => {
    stubTableConfigSheet({ 4: filledTableConfigRow });

    const sheet = fetchedTableConfig();
    sheet.topRow.clearValues();

    expect(
      sheet.topRow.valueOrEmpty(
        getColumnTraitByName("tableConfig", "letApiAccess", "columnId"),
      ),
    ).toBe(false);
    expect(sheet.topRow.isBlank).toBe(true);
  });

  // The default and the blank must agree, or an append and a read back disagree.
  it("defaults a checkbox cell to the same unchecked a blank reads as", () => {
    stubTableConfigSheet({ 4: filledTableConfigRow });

    const columnId = getColumnTraitByName(
      "tableConfig",
      "letApiAccess",
      "columnId",
    );
    const cell = fetchedTableConfig().topRow.cell(columnId);
    cell.updateToDefault();

    expect(cell.valueOrEmpty()).toBe(false);
  });

  it("defaults a date cell to blank rather than to today", () => {
    const dateColumnId = getColumnTraitByName(
      "valueTypes",
      "dateValue",
      "columnId",
    );
    const { grid } = stubSheetsService({
      sheets: [
        {
          sheetId: valueTypesGid,
          title: "Value Types",
          rows: buildGridRows({
            0: [valueTypesIdColumnId, dateColumnId],
            4: ["r:vty:row4", 45000],
          }),
          table: { endRowIndex: 5 },
        },
      ],
    });
    const ssi = new SpreadsheetIdentified(
      SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
    );
    const cell = ssi.tableOnSheet(valueTypesGid).topRow.cell(dateColumnId);
    cell.prepFetch();
    ssi.fetchAllPrepped();

    cell.updateToDefault();
    ssi.raw.batchUpdateGSheets();

    expect(grid.sheet(valueTypesGid).values(topRowRange)).toEqual([
      ["r:vty:row4", ""],
    ]);
    expect(cell.valueOrEmpty()).toBe("");
  });
});

describe("TableIdentified.appendRowDefault", () => {
  it("throws when the Table's one data row was never fetched, naming the Table and the prefetch owed", () => {
    stubTableConfigWithUnreadTopRow();

    expect(() => unfetchedTableConfig().appendRowDefault()).toThrowError(
      /Table ".*" on "Table Config".*never fetched.*Prefetch its top data row/,
    );
  });

  it("skips a configured column missing from the column-ID row and writes the rest", () => {
    const omittedColumnId = getColumnTraitByName(
      "tableConfig",
      "letApiAccess",
      "columnId",
    );
    const columnIdRow = tableConfigColumnIdRow.filter(
      (columnId) => columnId !== omittedColumnId,
    );
    const { grid } = stubSheetsService({
      sheets: [
        {
          sheetId: tableConfigGid,
          title: "Table Config",
          rows: buildGridRows({
            0: columnIdRow,
            4: [null, null, null],
          }),
          table: { endRowIndex: 5 },
        },
      ],
    });

    const ssi = new SpreadsheetIdentified(
      SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
    );
    const sheet = ssi.tableOnSheet(tableConfigGid);
    sheet.topRow.prepFetchFull();
    ssi.fetchAllPrepped();
    sheet.appendRowDefault();
    ssi.raw.batchUpdateGSheets();

    expect(
      grid.sheet(tableConfigGid).values({ ...topRowRange, endColumnIndex: 4 }),
    ).toEqual([["", "", "", null]]);
  });
});

describe("Identified formula writes", () => {
  beforeEach(() => {
    stubSheetsService({
      sheets: [
        {
          sheetId: computedGid,
          title: "Computed",
          rows: buildGridRows({ 0: [amountColumnId, rowNumberColumnId] }),
          table: { endRowIndex: 6 },
        },
      ],
    });
  });

  it("refuses a formula write on a non-formula column", () => {
    const ssi = new SpreadsheetIdentified(
      SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
    );
    ssi.raw.fetchAllSheetProperties();

    expect(() =>
      ssi
        .tableOnSheet(computedGid)
        .column(amountColumnId)
        .updateAllFormulas("=1"),
    ).toThrowError(/not a formula column/);
  });

  it("queues a formula write on a formula column", () => {
    const ssi = new SpreadsheetIdentified(
      SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
    );
    ssi.raw.fetchAllSheetProperties();

    expect(() =>
      ssi
        .tableOnSheet(computedGid)
        .column(rowNumberColumnId)
        .updateAllFormulas("=ROW()"),
    ).not.toThrow();
  });
  it("writes a formula column's head cell as plain text", () => {
    const ssi = new SpreadsheetIdentified(
      SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
    );
    ssi.raw.fetchAllSheetProperties();

    expect(() =>
      ssi
        .tableOnSheet(computedGid)
        .column(rowNumberColumnId)
        .headCell("header")
        .updateValue("Row number"),
    ).not.toThrow();
  });
});

describe("SpreadsheetIdentified.fetchAllPrepped / FetchTargetIdentified", () => {
  function recordedGridRanges(calls: object[]): unknown[] {
    const resource = calls[0] as {
      dataFilters: { gridRange: unknown }[];
    };
    return resource.dataFilters.map((filter) => filter.gridRange);
  }

  function valueTypesWithTwoColumns() {
    return stubSheetsService({
      sheets: [
        {
          sheetId: valueTypesGid,
          title: "Value Types",
          rows: buildGridRows({
            0: [valueTypesIdColumnId, checkboxColumnId],
            3: ["ID", "Checkbox"],
            4: ["r:vty:row4", true],
            5: [null, null],
          }),
          table: { endRowIndex: 6 },
        },
      ],
    });
  }

  function lastFetchedRanges(getByDataFilterCalls: object[]): unknown[] {
    const lastCall = getByDataFilterCalls[getByDataFilterCalls.length - 1];
    if (lastCall === undefined) return [];
    return recordedGridRanges([lastCall]);
  }

  it("narrows FetchTargetIdentified by kind", () => {
    type FullRow = Extract<FetchTargetIdentified, { kind: "fullRow" }>;
    type FullColumn = Extract<
      FetchTargetIdentified,
      { kind: "fullDataColumn" }
    >;
    type SingleCell = Extract<FetchTargetIdentified, { kind: "singleCell" }>;

    assertType<IsExactly<FullRow, { kind: "fullRow"; row: number }>>(true);
    assertType<
      IsExactly<FullColumn, { kind: "fullDataColumn"; column: string }>
    >(true);
    assertType<
      IsExactly<SingleCell, { kind: "singleCell"; row: number; column: string }>
    >(true);
  });

  it("resolves a full-row target to that row's table columns", () => {
    const { getByDataFilterCalls } = valueTypesWithTwoColumns();
    const ssi = new SpreadsheetIdentified(
      SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
    );
    ssi.tableOnSheet(valueTypesGid).topRow.prepFetchFull();
    ssi.fetchAllPrepped();

    expect(lastFetchedRanges(getByDataFilterCalls)).toEqual(
      expect.arrayContaining([
        {
          sheetId: valueTypesGid,
          startRowIndex: 4,
          endRowIndex: 5,
          startColumnIndex: 0,
          endColumnIndex: 2,
        },
      ]),
    );
  });

  it("resolves a full-data-column target to that column from the column ID row to the last row", () => {
    const { getByDataFilterCalls } = valueTypesWithTwoColumns();
    const ssi = new SpreadsheetIdentified(
      SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
    );
    ssi
      .tableOnSheet(valueTypesGid)
      .column(valueTypesIdColumnId)
      .prepFetchFull();
    ssi.fetchAllPrepped();

    expect(lastFetchedRanges(getByDataFilterCalls)).toEqual(
      expect.arrayContaining([
        {
          sheetId: valueTypesGid,
          startRowIndex: 0,
          endRowIndex: 6,
          startColumnIndex: 0,
          endColumnIndex: 1,
        },
      ]),
    );
  });

  it("resolves a single-cell target to that cell's grid range", () => {
    const { getByDataFilterCalls } = valueTypesWithTwoColumns();
    const ssi = new SpreadsheetIdentified(
      SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
    );
    ssi
      .tableOnSheet(valueTypesGid)
      .column(valueTypesIdColumnId)
      .cell(filledRowIndex)
      .prepFetch();
    ssi.fetchAllPrepped();

    expect(lastFetchedRanges(getByDataFilterCalls)).toEqual(
      expect.arrayContaining([
        {
          sheetId: valueTypesGid,
          startRowIndex: 4,
          endRowIndex: 5,
          startColumnIndex: 0,
          endColumnIndex: 1,
        },
      ]),
    );
  });
});

describe("SpreadsheetIdentified Tables", () => {
  const valueTypesTableId = getTableTraitByName("valueTypes", "tableId");

  function stubValueTypes(columnIdRow: string[], header: string[]) {
    const idAt = columnIdRow.indexOf(valueTypesIdColumnId);
    const bodyRow = columnIdRow.map((_, colIndex) =>
      colIndex === idAt ? "r:vty:row4" : true,
    );
    return stubSheetsService({
      sheets: [
        {
          sheetId: valueTypesGid,
          title: "Value Types",
          rows: buildGridRows({ 0: columnIdRow, 3: header, 4: bodyRow }),
          table: { endRowIndex: 5 },
        },
      ],
    });
  }

  function initIdentified(): SpreadsheetIdentified {
    return new SpreadsheetIdentified(
      SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
    );
  }

  it("resolves a column moved within its Table to its new index, with no config change", () => {
    stubValueTypes(
      [checkboxColumnId, valueTypesIdColumnId],
      ["Checkbox", "ID"],
    );
    const ssi = initIdentified();
    const column = ssi.tableOnSheet(valueTypesGid).column(valueTypesIdColumnId);
    column.prepFetchFull();
    ssi.fetchAllPrepped();

    expect(column.colIndex).toBe(1);
    expect(column.valueOrEmpty(0)).toBe("r:vty:row4");
  });

  it("stops on a lone Table inserted again under a new ID, offering regeneration, and sends nothing", () => {
    const { batchUpdateCount } = stubSheetsService({
      sheets: [
        {
          sheetId: valueTypesGid,
          title: "Value Types",
          rows: buildGridRows({
            0: [valueTypesIdColumnId],
            3: ["ID"],
            4: ["r:vty:row4"],
          }),
          table: { tableId: "recreated", endRowIndex: 5 },
        },
      ],
    });
    const ssi = initIdentified();
    ssi
      .tableOnSheet(valueTypesGid)
      .column(valueTypesIdColumnId)
      .prepFetchFull();

    expect(() => ssi.fetchAllPrepped()).toThrow(
      /"Value Types" \(gid \d+\) has no Table "valueTypes" .*regenerate the configs/,
    );
    expect(batchUpdateCount()).toBe(0);
  });

  it("addresses a known Table by its tableId", () => {
    stubValueTypes(
      [valueTypesIdColumnId, checkboxColumnId],
      ["ID", "Checkbox"],
    );
    const ssi = initIdentified();
    ssi.tableOnSheet(valueTypesGid).ensureColumnIdsAreFetched();
    const column = ssi.table(valueTypesTableId).column(valueTypesIdColumnId);
    column.prepFetchFull();
    ssi.fetchAllPrepped();

    expect(column.valueOrEmpty(0)).toBe("r:vty:row4");
  });

  it("reaches the recorded Table by its sheet when an unmanaged Table sits beside it", () => {
    const recordedTableId = getTableTraitByName("valueTypes", "tableId");
    stubSheetsService({
      sheets: [
        {
          sheetId: valueTypesGid,
          title: "Value Types",
          rows: buildGridRows({
            0: [valueTypesIdColumnId, checkboxColumnId],
            3: ["ID", "Checkbox"],
            4: ["r:vty:row4", true],
          }),
          tables: [
            {
              tableId: "unmanaged",
              startColumnIndex: 10,
              endColumnIndex: 11,
              endRowIndex: 5,
            },
            { tableId: recordedTableId, endColumnIndex: 2, endRowIndex: 5 },
          ],
        },
      ],
    });
    const ssi = initIdentified();
    const table = ssi.tableOnSheet(valueTypesGid).ensureColumnIdsAreFetched();
    const column = table.column(valueTypesIdColumnId);
    column.prepFetchFull();
    ssi.fetchAllPrepped();

    expect(table.tableId).toBe(recordedTableId);
    expect(column.valueOrEmpty(0)).toBe("r:vty:row4");
  });

  it("queues fetches per Table, whether reached through its sheet or its tableId", () => {
    stubValueTypes(
      [valueTypesIdColumnId, checkboxColumnId],
      ["ID", "Checkbox"],
    );
    const ssi = initIdentified();
    ssi.tableOnSheet(valueTypesGid).ensureColumnIdsAreFetched();
    ssi
      .tableOnSheet(valueTypesGid)
      .column(valueTypesIdColumnId)
      .prepFetchFull();
    ssi.table(valueTypesTableId).column(checkboxColumnId).prepFetchFull();

    expect(
      ssi.tablesPreppedForFetch.map((table) => table.fetchTargets),
    ).toEqual([
      [
        { kind: "fullDataColumn", column: valueTypesIdColumnId },
        { kind: "fullDataColumn", column: checkboxColumnId },
      ],
    ]);
  });

  it("keeps a fetch prepped through the sheet before its properties arrive", () => {
    stubValueTypes(
      [valueTypesIdColumnId, checkboxColumnId],
      ["ID", "Checkbox"],
    );
    const ssi = initIdentified();
    ssi.tableOnSheet(valueTypesGid).column(checkboxColumnId).prepFetchFull();
    ssi.tableOnSheet(valueTypesGid).ensureColumnIdsAreFetched();
    ssi.table(valueTypesTableId).column(valueTypesIdColumnId).prepFetchFull();
    ssi.fetchAllPrepped();

    const table = ssi.table(valueTypesTableId);
    expect(table.column(valueTypesIdColumnId).valueOrEmpty(0)).toBe(
      "r:vty:row4",
    );
    expect(table.column(checkboxColumnId).valueOrEmpty(0)).toBe(true);
    expect(ssi.tablesPreppedForFetch).toEqual([]);
  });

  it("fetches a sheet's rules and protections with its Table's column IDs", () => {
    const { getByDataFilterCalls } = stubSheetsService({
      sheets: [
        {
          sheetId: valueTypesGid,
          title: "Value Types",
          rows: buildGridRows({
            0: [valueTypesIdColumnId],
            3: ["ID"],
            4: ["r:vty:row4"],
          }),
          table: { endRowIndex: 5 },
          conditionalFormats: [],
          protectedRanges: [],
        },
      ],
    });
    const ssi = initIdentified();
    const table = ssi.tableOnSheet(valueTypesGid);
    table.raw.sheet
      .gatherFetchConditionalFormatRules()
      .gatherFetchEditProtections();
    ssi.fetchAllPrepped();
    const fetchCount = getByDataFilterCalls.length;

    expect(table.raw.sheet.conditionalFormatRules()).toEqual([]);
    expect(table.raw.sheet.editProtections()).toEqual([]);
    expect(table.column(valueTypesIdColumnId).colIndex).toBe(0);
    expect(getByDataFilterCalls).toHaveLength(fetchCount);
  });
});

describe("SpreadsheetIdentified.fetchAllGathered, the placement check", () => {
  function initIdentified(): SpreadsheetIdentified {
    return new SpreadsheetIdentified(
      SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
    );
  }

  it("stops naming every managed sheet whose header zone found no Table, and says to regenerate", () => {
    const { batchUpdateCount } = stubSheetsService({
      sheets: [
        { sheetId: itemGid, title: "Item" },
        { sheetId: logGid, title: "Log" },
      ],
    });

    const ssi = initIdentified();
    ssi.raw.sheet(itemGid).table(itemTableId).gatherFetchProperties();
    ssi.raw.table(itemTableId).columnResolver.gatherFetchColumnIds();
    ssi.raw.sheet(logGid).table(logTableId).gatherFetchProperties();

    expect(() => ssi.fetchAllGathered()).toThrowError(
      /"Item" \(gid \d+\) has no Table "item" with its header row on row 4 — move it back, or regenerate the configs.*"Log" \(gid \d+\) has no Table/,
    );
    expect(batchUpdateCount()).toBe(0);
  });

  it("stops on a Table one row too high for its head rows to fit, naming the header row it needs", () => {
    const { batchUpdateCount } = stubSheetsService({
      sheets: [
        misplacedTableSheet({
          sheetId: itemGid,
          title: "Item",
          startRowIndex: tableHeaderRowIndex - 1,
        }),
      ],
    });

    const ssi = initIdentified();
    ssi.raw.sheet(itemGid).table(itemTableId).gatherFetchProperties();

    expect(() => ssi.fetchAllGathered()).toThrowError(
      /"Item" \(gid \d+\) has Table ".*", which must have its header row on row 4 — move it back$/,
    );
    expect(batchUpdateCount()).toBe(0);
  });

  it("judges a Table that fetches its own column IDs", () => {
    const { batchUpdateCount } = stubSheetsService({
      sheets: [
        misplacedTableSheet({
          sheetId: itemGid,
          title: "Item",
          startRowIndex: tableHeaderRowIndex - 1,
        }),
      ],
    });

    const ssi = initIdentified();

    expect(() =>
      ssi.tableOnSheet(itemGid).ensureColumnIdsAreFetched(),
    ).toThrowError(/"Item" \(gid \d+\) has Table ".*", which must have/);
    expect(batchUpdateCount()).toBe(0);
  });

  it("stops a run using a header-only managed Table with its fix, sending nothing", () => {
    const { batchUpdateCount } = stubSheetsService({
      sheets: [
        {
          ...placedTableSheet({ sheetId: logGid, title: "Log" }),
          table: { endRowIndex: topDataRowIndex, name: "Logs" },
        },
      ],
    });

    const ssi = initIdentified();
    ssi.raw.sheet(logGid).table(logTableId).gatherFetchProperties();

    expect(thrownMessage(() => ssi.fetchAllGathered())).toMatch(
      /^Table "Logs" on "Log" \(gid \d+\) has only its header: add a row below it holding its formulas\.$/,
    );
    expect(batchUpdateCount()).toBe(0);
  });

  it("names the header-only Table alone when a misplaced one is used beside it", () => {
    const { batchUpdateCount } = stubSheetsService({
      sheets: [
        misplacedTableSheet({
          sheetId: itemGid,
          title: "Item",
          startRowIndex: tableHeaderRowIndex - 1,
        }),
        {
          ...placedTableSheet({ sheetId: logGid, title: "Log" }),
          table: { endRowIndex: topDataRowIndex, name: "Logs" },
        },
      ],
    });

    const ssi = initIdentified();
    ssi.raw.sheet(itemGid).table(itemTableId).gatherFetchProperties();
    ssi.raw.sheet(logGid).table(logTableId).gatherFetchProperties();

    expect(thrownMessage(() => ssi.fetchAllGathered())).toMatch(
      /^Table "Logs" on "Log" \(gid \d+\) has only its header: add a row below it holding its formulas\.$/,
    );
    expect(batchUpdateCount()).toBe(0);
  });

  function stubItemAndLogMisplaced(): { batchUpdateCount(): number } {
    return stubSheetsService({
      isEveryTableInFilteredFetch: true,
      sheets: [
        misplacedTableSheet({
          sheetId: itemGid,
          title: "Item",
          startRowIndex: tableHeaderRowIndex - 1,
        }),
        misplacedTableSheet({
          sheetId: logGid,
          title: "Log",
          startRowIndex: tableHeaderRowIndex - 2,
        }),
      ],
    });
  }

  it("names every misplaced sheet the run gathered for in one error", () => {
    const { batchUpdateCount } = stubItemAndLogMisplaced();

    const ssi = initIdentified();
    ssi.raw.sheet(itemGid).table(itemTableId).gatherFetchProperties();
    ssi.raw.sheet(logGid).table(logTableId).gatherFetchProperties();

    expect(() => ssi.fetchAllGathered()).toThrowError(/"Item".*"Log"/);
    expect(batchUpdateCount()).toBe(0);
  });

  it("leaves a misplaced Table the run did not gather for unnamed", () => {
    const { batchUpdateCount } = stubItemAndLogMisplaced();

    const ssi = initIdentified();
    ssi.raw.sheet(itemGid).table(itemTableId).gatherFetchProperties();

    const message = thrownMessage(() => ssi.fetchAllGathered());
    expect(message).toMatch(/"Item"/);
    expect(message).not.toMatch(/"Log"/);
    expect(batchUpdateCount()).toBe(0);
  });

  it("stops on a Table moved below the header zone as missing, in the one round trip", () => {
    const { getByDataFilterCalls, getCalls, batchUpdateCount } =
      stubSheetsService({
        sheets: [
          misplacedTableSheet({
            sheetId: itemGid,
            title: "Item",
            startRowIndex: tableHeaderRowIndex + 2,
          }),
        ],
      });

    const ssi = initIdentified();
    ssi.raw.sheet(itemGid).table(itemTableId).gatherFetchProperties();
    ssi.raw.table(itemTableId).columnResolver.gatherFetchColumnIds();

    expect(() => ssi.fetchAllGathered()).toThrowError(
      /"Item" \(gid \d+\) has no Table "item" with its header row on row 4 — move it back, or regenerate the configs/,
    );
    expect(getByDataFilterCalls).toHaveLength(1);
    expect(getCalls).toEqual([]);
    expect(batchUpdateCount()).toBe(0);
  });

  it("stops on a head band shifted down by an inserted row while the header stayed put", () => {
    const { batchUpdateCount } = stubSheetsService({
      sheets: [
        {
          ...placedTableSheet({ sheetId: itemGid, title: "Item" }),
          rows: buildGridRows({
            [colIdRowIndex + 1]: [ownColumnId(itemGid)],
            [tableHeaderRowIndex]: ["ID"],
          }),
        },
      ],
    });

    const ssi = initIdentified();
    ssi.raw.sheet(itemGid).table(itemTableId).gatherFetchProperties();

    expect(() => ssi.fetchAllGathered()).toThrowError(
      /"Item" \(gid \d+\) needs its own "itm" column IDs, and only those, in row 1 — move the Table back, or regenerate the configs/,
    );
    expect(batchUpdateCount()).toBe(0);
  });

  it("stops on another Table's column ID in the column ID row", () => {
    const { batchUpdateCount } = stubSheetsService({
      sheets: [
        {
          ...placedTableSheet({ sheetId: itemGid, title: "Item" }),
          rows: buildGridRows({
            [colIdRowIndex]: [ownColumnId(itemGid), ownColumnId(logGid)],
            [tableHeaderRowIndex]: ["ID", "Name"],
          }),
          table: { endRowIndex: tableEndRowIndex, endColumnIndex: 2 },
        },
      ],
    });

    const ssi = initIdentified();
    ssi.raw.sheet(itemGid).table(itemTableId).gatherFetchProperties();
    ssi.raw.table(itemTableId).columnResolver.gatherFetchColumnIds();

    expect(() => ssi.fetchAllGathered()).toThrowError(
      /"Item" \(gid \d+\) needs its own "itm" column IDs, and only those, in row 1/,
    );
    expect(batchUpdateCount()).toBe(0);
  });

  it("takes a recorded Table as missing when its sheet holds several Tables and none carries its ID", () => {
    const { batchUpdateCount } = stubSheetsService({
      isEveryTableInFilteredFetch: true,
      sheets: [extraTablesSheet({ sheetId: itemGid, title: "Item" })],
    });

    const ssi = initIdentified();
    ssi.raw.sheet(itemGid).table(itemTableId).gatherFetchProperties();

    expect(() => ssi.fetchAllGathered()).toThrowError(
      /"Item" \(gid \d+\) has no Table "item" with its header row on row 4/,
    );
    expect(batchUpdateCount()).toBe(0);
  });

  it("sends the header zone once on a fetch after one that stopped on placement", () => {
    const { getByDataFilterCalls } = stubSheetsService({
      sheets: [
        misplacedTableSheet({
          sheetId: itemGid,
          title: "Item",
          startRowIndex: tableHeaderRowIndex - 1,
        }),
      ],
    });

    const ssi = initIdentified();
    ssi.raw.sheet(itemGid).table(itemTableId).gatherFetchProperties();
    expect(() => ssi.fetchAllGathered()).toThrowError(/must have/);
    ssi.raw.table(itemTableId).gatherFetchProperties();
    expect(() => ssi.fetchAllGathered()).toThrowError(/must have/);

    expect(recordedGridRanges(getByDataFilterCalls.slice(1))).toEqual([
      {
        sheetId: itemGid,
        startRowIndex: 0,
        endRowIndex: tableHeaderRowIndex + 1,
      },
    ]);
  });

  describe("several managed Tables on one sheet", () => {
    const rightId = layoutTableId("layoutRight");

    function gatherEveryZone(ssi: SpreadsheetIdentified): void {
      managedLayoutTableNames.forEach((tableName) => {
        const table = ssi.raw.sheet(layoutGid).table(layoutTableId(tableName));
        table.gatherFetchProperties();
        table.columnResolver.gatherFetchColumnIds();
      });
    }

    it("names a Table met below the header zone while its neighbours pass", () => {
      const { batchUpdateCount } = stubSheetsService({
        isEveryTableInFilteredFetch: true,
        sheets: [layoutSheet({ layoutRight: { startRowIndex: 5 } })],
      });

      const ssi = initIdentified();
      gatherEveryZone(ssi);

      expect(thrownMessage(() => ssi.fetchAllGathered())).toMatch(
        /^1 managed Table\(s\) are missing or misplaced: "Layout" \(gid \d+\) has Table "layoutRight", which must have its header row on row 4 — move it back$/,
      );
      expect(batchUpdateCount()).toBe(0);
    });

    it("names a Table missing from the sheet it shares", () => {
      const sheet = layoutSheet();
      const { batchUpdateCount } = stubSheetsService({
        sheets: [
          {
            ...sheet,
            tables: sheet.tables?.filter((table) => table.tableId !== rightId),
          },
        ],
      });

      const ssi = initIdentified();
      gatherEveryZone(ssi);

      expect(() => ssi.fetchAllGathered()).toThrowError(
        /^1 managed Table\(s\) .*"Layout" \(gid \d+\) has no Table "layoutRight" with its header row on row 4 — move it back, or regenerate the configs with sheets-framework gen-configs if it is gone$/,
      );
      expect(batchUpdateCount()).toBe(0);
    });

    it("names a Table whose column ID row holds another Table's IDs", () => {
      const { batchUpdateCount } = stubSheetsService({
        sheets: [
          layoutSheet({
            layoutRight: {
              headRows: { 3: ["c:lyl:entry", "c:lyl:amount"] },
            },
          }),
        ],
      });

      const ssi = initIdentified();
      gatherEveryZone(ssi);

      expect(() => ssi.fetchAllGathered()).toThrowError(
        /^1 managed Table\(s\) .*"Layout" \(gid \d+\) needs its own "lyr" column IDs, and only those, in row 1 — move the Table back, or regenerate the configs with sheets-framework gen-configs$/,
      );
      expect(batchUpdateCount()).toBe(0);
    });
  });
});
