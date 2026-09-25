import { beforeEach, describe, expect, it } from "vitest";

import { getColumnTraitByName } from "../01_SpreadsheetSchema/columnConfigsTypes";
import { getSheetTraitByName } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import type { Value, VnToCvn } from "../01_SpreadsheetSchema/valueSchemas";
import {
  blankSheetConfigRow,
  filledSheetConfigRow,
  sheetConfigColumnIdRow,
  sheetConfigGid,
  stubSheetConfigSheet,
} from "../testSupport/fakeSheetConfigSheet";
import {
  buildGridRows,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { assertType, type IsExactly } from "../testSupport/typeAssertions";
import { SpreadsheetBaseIdentified } from "./ClassBases/SpreadsheetBaseIdentified";
import type { FetchTargetIdentified } from "./ClassTypes/StateIdentified";
import { ColumnIdentified } from "./ColumnIdentified";
import { ColumnMetaIdentified } from "./ColumnMetaIdentified";
import { RowIdentified } from "./RowIdentified";
import { SheetIdentified } from "./SheetIdentified";
import { SheetMetaIdentified } from "./SheetMetaIdentified";
import { SpreadsheetIdentified } from "./SpreadsheetIdentified";

const itemGid = getSheetTraitByName("item", "sheetGid");
const itemIdColumnId = getColumnTraitByName("item", "id", "columnId");

// A mis-wired accessor still type-checks; the instance checks catch it.
describe("SpreadsheetIdentified navigation", () => {
  it("gives each accessor the class its return type names", () => {
    stubSheetsService();
    const ssi = new SpreadsheetIdentified(
      SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
    );
    const sheet = ssi.sheet(itemGid);
    const sheetMeta = ssi.sheetMeta(itemGid);
    const column = sheet.column(itemIdColumnId);
    const columnMeta = sheetMeta.column(itemIdColumnId);

    assertType<IsExactly<typeof sheet, SheetIdentified>>(true);
    assertType<IsExactly<typeof sheetMeta, SheetMetaIdentified>>(true);
    assertType<IsExactly<typeof sheet.meta, SheetMetaIdentified>>(true);
    assertType<IsExactly<typeof sheetMeta.primary, SheetIdentified>>(true);
    assertType<IsExactly<typeof column, ColumnIdentified>>(true);
    assertType<IsExactly<typeof columnMeta, ColumnMetaIdentified>>(true);
    assertType<IsExactly<typeof column.sheet, SheetIdentified>>(true);
    assertType<IsExactly<typeof columnMeta.sheet, SheetMetaIdentified>>(true);
    assertType<IsExactly<typeof column.meta, ColumnMetaIdentified>>(true);
    assertType<IsExactly<typeof columnMeta.primary, ColumnIdentified>>(true);
    assertType<IsExactly<ReturnType<typeof sheet.row>, RowIdentified>>(true);

    expect(sheet.meta).toBeInstanceOf(SheetMetaIdentified);
    expect(sheetMeta.primary).toBeInstanceOf(SheetIdentified);
    expect(column).toBeInstanceOf(ColumnIdentified);
    expect(columnMeta).toBeInstanceOf(ColumnMetaIdentified);
    expect(column.sheet).toBeInstanceOf(SheetIdentified);
    expect(columnMeta.sheet).toBeInstanceOf(SheetMetaIdentified);
    expect(column.meta).toBeInstanceOf(ColumnMetaIdentified);
    expect(columnMeta.primary).toBeInstanceOf(ColumnIdentified);
    expect(sheet.row(sheet.schema.topDataRowIdx)).toBeInstanceOf(RowIdentified);
  });
});

const valueTypesGid = getSheetTraitByName("valueTypes", "sheetGid");
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
const filledRowIndex = 4;
const blankRowIndex = 5;

// Row 5 is the blank row; its checkbox is untouched, so it reads blank not false.
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

function fetchedValueTypesSheet(): SheetIdentified {
  const ssi = new SpreadsheetIdentified(
    SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
  );
  const sheet = ssi.sheet(valueTypesGid);
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
      new RegExp(`${valueTypesIdColumnId}.*${blankRowIndex}`),
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
          rowsWithNoGridData: [blankRowIndex],
          table: { endRowIndex: 6 },
        },
      ],
    });
    const ssi = new SpreadsheetIdentified(
      SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
    );
    const sheet = ssi.sheet(valueTypesGid);
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
      ...sheet.sheetIdentifiedProps,
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
function stubSheetConfigWithUnreadTopRow() {
  return stubSheetConfigSheet({ 4: blankSheetConfigRow }, [4]);
}

function fetchedSheetConfig(): SheetIdentified {
  const ssi = new SpreadsheetIdentified(
    SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
  );
  const sheet = ssi.sheet(sheetConfigGid);
  sheet.topRow.prepFetchFull();
  ssi.fetchAllPrepped();
  return sheet;
}

function unfetchedSheetConfig(): SheetIdentified {
  const ssi = new SpreadsheetIdentified(
    SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
  );
  ssi.sheetMeta(sheetConfigGid).ensureColumnIdsAreFetched();
  return ssi.sheet(sheetConfigGid);
}

describe("RowIdentified.isBlank / isReusable", () => {
  it("calls a row whose every non-formula cell is empty blank", () => {
    stubSheetConfigSheet({ 4: blankSheetConfigRow });

    expect(fetchedSheetConfig().topRow.isBlank).toBe(true);
  });

  it("calls a row holding any non-formula value not blank", () => {
    stubSheetConfigSheet({ 4: filledSheetConfigRow });

    expect(fetchedSheetConfig().topRow.isBlank).toBe(false);
  });

  it("calls a row nothing fetched not blank, since nothing read it", () => {
    stubSheetConfigWithUnreadTopRow();

    expect(unfetchedSheetConfig().topRow.isBlank).toBe(false);
  });

  it("makes a blank row reusable until an append reserves it", () => {
    stubSheetConfigSheet({ 4: blankSheetConfigRow });

    const sheet = fetchedSheetConfig();
    expect(sheet.topRow.isReusable).toBe(true);

    sheet.appendRowDefault();
    expect(sheet.topRow.isReusable).toBe(false);
  });
});

describe("SheetIdentified.hasNoData", () => {
  it("is true for a sheet whose one row is blank", () => {
    stubSheetConfigSheet({ 4: blankSheetConfigRow });

    expect(fetchedSheetConfig().hasNoData).toBe(true);
  });

  it("is false while any row still holds data", () => {
    stubSheetConfigSheet({ 4: blankSheetConfigRow, 5: filledSheetConfigRow });

    expect(fetchedSheetConfig().hasNoData).toBe(false);
  });
});

describe("RowIdentified.clearValues", () => {
  it("empties every non-formula cell and touches no formula cell", () => {
    const { batchUpdateCalls } = stubSheetConfigSheet({
      4: filledSheetConfigRow,
    });

    const ssi = new SpreadsheetIdentified(
      SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
    );
    const sheet = ssi.sheet(sheetConfigGid);
    sheet.topRow.prepFetchFull();
    ssi.fetchAllPrepped();
    sheet.topRow.clearValues();
    ssi.raw.batchUpdateGSheets();

    expect(writtenValuesByColIndex(batchUpdateCalls)).toEqual([
      [0, ""],
      [1, ""],
      [2, ""],
    ]);
    expect(sheet.topRow.isBlank).toBe(true);
  });

  // The cell is cleared to a blank on the wire; the value name is what reads it back.
  it("leaves a cleared checkbox reading unchecked rather than blank", () => {
    stubSheetConfigSheet({ 4: filledSheetConfigRow });

    const sheet = fetchedSheetConfig();
    sheet.topRow.clearValues();

    expect(
      sheet.topRow.valueOrEmpty(
        getColumnTraitByName("sheetConfig", "letApiAccess", "columnId"),
      ),
    ).toBe(false);
    expect(sheet.topRow.isBlank).toBe(true);
  });

  // The default and the blank must agree, or an append and a read back disagree.
  it("defaults a checkbox cell to the same unchecked a blank reads as", () => {
    stubSheetConfigSheet({ 4: filledSheetConfigRow });

    const columnId = getColumnTraitByName(
      "sheetConfig",
      "letApiAccess",
      "columnId",
    );
    const cell = fetchedSheetConfig().topRow.cell(columnId);
    cell.updateToDefault();

    expect(cell.valueOrEmpty()).toBe(false);
  });

  it("defaults a date cell to blank rather than to today", () => {
    const dateColumnId = getColumnTraitByName(
      "valueTypes",
      "dateValue",
      "columnId",
    );
    const { batchUpdateCalls } = stubSheetsService({
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
    const cell = ssi.sheet(valueTypesGid).topRow.cell(dateColumnId);
    cell.prepFetch();
    ssi.fetchAllPrepped();

    cell.updateToDefault();
    ssi.raw.batchUpdateGSheets();

    expect(writtenValuesByColIndex(batchUpdateCalls)).toEqual([[1, ""]]);
    expect(cell.valueOrEmpty()).toBe("");
  });
});

describe("SheetIdentified.appendRowDefault", () => {
  it("throws when the sheet's one data row was never fetched, naming the prefetch owed", () => {
    stubSheetConfigWithUnreadTopRow();

    expect(() => unfetchedSheetConfig().appendRowDefault()).toThrowError(
      /never fetched.*Prefetch that row first/,
    );
  });

  it("skips a configured column missing from the column-ID row and writes the rest", () => {
    const omittedColumnId = getColumnTraitByName(
      "sheetConfig",
      "letApiAccess",
      "columnId",
    );
    const columnIdRow = sheetConfigColumnIdRow.filter(
      (columnId) => columnId !== omittedColumnId,
    );
    const { batchUpdateCalls } = stubSheetsService({
      sheets: [
        {
          sheetId: sheetConfigGid,
          title: "Sheet Config",
          rows: buildGridRows({
            0: columnIdRow,
            4: [null, null],
          }),
          table: { endRowIndex: 5 },
        },
      ],
    });

    const ssi = new SpreadsheetIdentified(
      SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
    );
    const sheet = ssi.sheet(sheetConfigGid);
    sheet.topRow.prepFetchFull();
    ssi.fetchAllPrepped();
    sheet.appendRowDefault();
    ssi.raw.batchUpdateGSheets();

    expect(writtenValuesByColIndex(batchUpdateCalls)).toEqual([
      [0, ""],
      [1, ""],
    ]);
  });
});

const computedGid = getSheetTraitByName("computed", "sheetGid");
const amountColumnId = getColumnTraitByName("computed", "amount", "columnId");
const rowNumberColumnId = getColumnTraitByName(
  "computed",
  "rowNumber",
  "columnId",
);

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
      ssi.sheet(computedGid).column(amountColumnId).updateAllFormulas("=1"),
    ).toThrowError(/not a formula column/);
  });

  it("queues a formula write on a formula column", () => {
    const ssi = new SpreadsheetIdentified(
      SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
    );
    ssi.raw.fetchAllSheetProperties();

    expect(() =>
      ssi
        .sheet(computedGid)
        .column(rowNumberColumnId)
        .updateAllFormulas("=ROW()"),
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
    ssi.sheet(valueTypesGid).topRow.prepFetchFull();
    ssi.fetchAllPrepped();

    expect(lastFetchedRanges(getByDataFilterCalls)).toEqual(
      expect.arrayContaining([
        {
          sheetId: valueTypesGid,
          startRowIndex: filledRowIndex,
          endRowIndex: filledRowIndex + 1,
          startColumnIndex: 0,
        },
      ]),
    );
  });

  it("resolves a full-data-column target to that column's data rows", () => {
    const { getByDataFilterCalls } = valueTypesWithTwoColumns();
    const ssi = new SpreadsheetIdentified(
      SpreadsheetBaseIdentified.initSpreadsheetIdentifiedProps(),
    );
    ssi.sheet(valueTypesGid).column(valueTypesIdColumnId).prepFetchFull();
    ssi.fetchAllPrepped();

    expect(lastFetchedRanges(getByDataFilterCalls)).toEqual(
      expect.arrayContaining([
        {
          sheetId: valueTypesGid,
          startRowIndex: filledRowIndex,
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
      .sheet(valueTypesGid)
      .column(valueTypesIdColumnId)
      .cell(filledRowIndex)
      .prepFetch();
    ssi.fetchAllPrepped();

    expect(lastFetchedRanges(getByDataFilterCalls)).toEqual(
      expect.arrayContaining([
        {
          sheetId: valueTypesGid,
          startRowIndex: filledRowIndex,
          endRowIndex: filledRowIndex + 1,
          startColumnIndex: 0,
          endColumnIndex: 1,
        },
      ]),
    );
  });
});

function writtenValuesByColIndex(
  calls: GoogleAppsScript.Sheets.Schema.BatchUpdateSpreadsheetRequest[],
): [number | undefined, unknown][] {
  return calls
    .flatMap((call) => call.requests ?? [])
    .filter((request) => request.updateCells)
    .map((request) => {
      const userEnteredValue =
        request.updateCells?.rows?.[0]?.values?.[0]?.userEnteredValue;
      return [
        request.updateCells?.range?.startColumnIndex,
        userEnteredValue?.stringValue ??
          userEnteredValue?.boolValue ??
          userEnteredValue?.numberValue,
      ];
    });
}
