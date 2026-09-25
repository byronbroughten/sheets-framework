import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import {
  type ColumnIsFormula,
  getColumnTraitByName,
} from "../01_SpreadsheetSchema/columnConfigsTypes";
import {
  getSheetTraitByName,
  type SheetName,
} from "../01_SpreadsheetSchema/sheetConfigsTypes";
import { ssConfigGet } from "../01_SpreadsheetSchema/spreadsheetConfigTypes";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  blankSheetConfigRow,
  filledSheetConfigRow,
  stubSheetConfigSheet,
} from "../testSupport/fakeSheetConfigSheet";
import {
  buildGridRows,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import {
  assertNotType,
  assertType,
  type IsExactly,
} from "../testSupport/typeAssertions";
import { SerialDate } from "../utils/SerialDate";
import type { SpreadsheetNamedProps } from "./ClassBases/SpreadsheetBaseNamed";
import { ColumnMetaNamed } from "./ColumnMetaNamed";
import { ColumnNamed } from "./ColumnNamed";
import { RowNamed } from "./RowNamed";
import { SheetMetaNamed } from "./SheetMetaNamed";
import { SheetNamed } from "./SheetNamed";
import { SpreadsheetNamed } from "./SpreadsheetNamed";

describe("SpreadsheetNamed props", () => {
  it("have no Named-state member", () => {
    type HasNamedState = "namedState" extends keyof SpreadsheetNamedProps
      ? true
      : false;
    assertType<IsExactly<HasNamedState, false>>(true);
  });
});

// A mis-wired accessor still type-checks; the instance checks catch it.
describe("SpreadsheetNamed dates", () => {
  const march14 = SerialDate.fromYmd({ year: 2024, month: 3, day: 14 });
  const march15 = SerialDate.fromYmd({ year: 2024, month: 3, day: 15 });

  beforeEach(() => {
    stubLogger();
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2024-03-15T02:30:00Z"));
  });
  afterEach(() => {
    vi.useRealTimers();
  });

  it("dates today in the fixture's default zone", () => {
    stubSheetsService();

    expect(SpreadsheetNamed.init().today()).toBe(march14);
  });

  it("dates today, now and serialDate.today in the spreadsheet's own zone", () => {
    stubSheetsService({ timeZone: "Asia/Tokyo" });
    const ss = SpreadsheetNamed.init();

    expect(ss.today()).toBe(march15);
    expect(ss.now()).toBe("2024-03-15 11:30:00");
    expect(ss.serialDate.today()).toBe(march15);
  });

  it("keeps the static date operations on serialDate", () => {
    stubSheetsService({ timeZone: "Asia/Tokyo" });

    expect(SpreadsheetNamed.init().serialDate.addDays(march14, 1)).toBe(
      march15,
    );
  });
});

describe("SpreadsheetNamed navigation", () => {
  it("gives each accessor the class its return type names", () => {
    stubSheetsService();
    const ss = SpreadsheetNamed.init();
    const sheet = ss.sheet("item");
    const sheetMeta = ss.sheetMeta("item");
    const column = sheet.column("id");
    const columnMeta = sheetMeta.column("id");

    assertType<IsExactly<typeof sheet, SheetNamed<"item">>>(true);
    assertType<IsExactly<typeof sheetMeta, SheetMetaNamed<"item">>>(true);
    assertType<IsExactly<typeof sheet.meta, SheetMetaNamed<"item">>>(true);
    assertType<IsExactly<typeof sheetMeta.primary, SheetNamed<"item">>>(true);
    assertType<IsExactly<typeof column, ColumnNamed<"item", "id">>>(true);
    assertType<IsExactly<typeof columnMeta, ColumnMetaNamed<"item", "id">>>(
      true,
    );
    assertType<IsExactly<typeof column.sheet, SheetNamed<"item">>>(true);
    assertType<IsExactly<typeof columnMeta.sheet, SheetMetaNamed<"item">>>(
      true,
    );
    assertType<IsExactly<typeof column.meta, ColumnMetaNamed<"item", "id">>>(
      true,
    );
    assertType<IsExactly<typeof columnMeta.primary, ColumnNamed<"item", "id">>>(
      true,
    );
    assertType<IsExactly<ReturnType<typeof sheet.row>, RowNamed<"item">>>(true);

    expect(sheet.meta).toBeInstanceOf(SheetMetaNamed);
    expect(sheetMeta.primary).toBeInstanceOf(SheetNamed);
    expect(column).toBeInstanceOf(ColumnNamed);
    expect(columnMeta).toBeInstanceOf(ColumnMetaNamed);
    expect(column.sheet).toBeInstanceOf(SheetNamed);
    expect(columnMeta.sheet).toBeInstanceOf(SheetMetaNamed);
    expect(column.meta).toBeInstanceOf(ColumnMetaNamed);
    expect(columnMeta.primary).toBeInstanceOf(ColumnNamed);
    expect(sheet.row(sheet.schema.topDataRowIdx)).toBeInstanceOf(RowNamed);
  });
});

const topDataRowIndex = ssConfigGet("tableHeaderRowIndexBase0") + 1;
const datesGid = getSheetTraitByName("dates", "sheetGid");
const valueTypesGid = getSheetTraitByName("valueTypes", "sheetGid");
const idColumnId = getColumnTraitByName("dates", "id", "columnId");
const requiredDateColumnId = getColumnTraitByName(
  "dates",
  "requiredDate",
  "columnId",
);
const optionalDateColumnId = getColumnTraitByName(
  "dates",
  "optionalDate",
  "columnId",
);
const requiredDateSerial = 45000;
const optionalDateSerial = 45365;
const filledRowIndex = 4;
const blankRowIndex = 5;

// Row 5 is the blank row; the checkbox is untouched, so it reads blank not false.
function stubDatesAndValueTypesWithBlankRow() {
  return stubSheetsService({
    sheets: [
      {
        sheetId: datesGid,
        title: "Dates",
        rows: buildGridRows({
          0: [idColumnId, requiredDateColumnId, optionalDateColumnId],
          3: ["ID", "Required date", "Optional date"],
          4: ["r:dat:row4", requiredDateSerial, optionalDateSerial],
          5: [null, null, null],
        }),
        table: { endRowIndex: 6 },
      },
      {
        sheetId: valueTypesGid,
        title: "Value Types",
        rows: buildGridRows({
          0: [
            getColumnTraitByName("valueTypes", "checkbox", "columnId"),
            getColumnTraitByName("valueTypes", "numberValue", "columnId"),
          ],
          3: ["Checkbox", "Number value"],
          4: [true, 7],
          5: [null, null],
        }),
        table: { endRowIndex: 6 },
      },
    ],
  });
}

function fetchedDatesSheet(): SheetNamed<"dates"> {
  const ss = SpreadsheetNamed.init();
  ss.sheet("dates").prepFetchColumnsFull("id", "requiredDate", "optionalDate");
  ss.fetchAllPrepped();
  return ss.sheet("dates");
}

function fetchedValueTypesSheet(): SheetNamed<"valueTypes"> {
  const ss = SpreadsheetNamed.init();
  ss.sheet("valueTypes").prepFetchColumnsFull("checkbox", "numberValue");
  ss.fetchAllPrepped();
  return ss.sheet("valueTypes");
}

describe("Named value accessors", () => {
  beforeEach(() => {
    stubDatesAndValueTypesWithBlankRow();
  });

  it("names the sheet, the column and the row when CellNamed.value hits a blank cell", () => {
    const cell = fetchedDatesSheet().column("id").cell(blankRowIndex);

    expect(() => cell.value()).toThrowError(
      new RegExp(`"id".*"dates".*${blankRowIndex}`),
    );
  });

  it("keeps the generated column id out of the Named message", () => {
    const cell = fetchedDatesSheet().column("id").cell(blankRowIndex);

    expect(() => cell.value()).not.toThrowError(new RegExp(idColumnId));
  });

  it("returns the empty string from CellNamed.valueOrEmpty on that same cell", () => {
    const cell = fetchedDatesSheet().column("id").cell(blankRowIndex);

    expect(cell.valueOrEmpty()).toBe("");
  });

  it("reads a filled cell identically through both forms", () => {
    const cell = fetchedDatesSheet().column("id").cell(filledRowIndex);

    expect(cell.value()).toBe("r:dat:row4");
    expect(cell.valueOrEmpty()).toBe("r:dat:row4");
  });

  // The same read one tier up, which is what proves Identified and Named agree.
  it("reads an untouched checkbox as unchecked, with no blank in the type", () => {
    const sheet = fetchedValueTypesSheet();
    const column = sheet.column("checkbox");

    expect(column.valueOrEmpty(blankRowIndex)).toBe(false);
    expect(column.value(blankRowIndex)).toBe(false);
    expect(column.value(filledRowIndex)).toBe(true);
    expect(sheet.row(blankRowIndex).value("checkbox")).toBe(false);
    assertType<IsExactly<ReturnType<typeof column.value>, boolean>>(true);
    assertType<IsExactly<ReturnType<typeof column.valueOrEmpty>, boolean>>(
      true,
    );
  });

  it("throws from ColumnNamed.value and returns empty from valueOrEmpty", () => {
    const column = fetchedDatesSheet().column("id");

    expect(() => column.value(blankRowIndex)).toThrowError(/is empty/);
    expect(column.valueOrEmpty(blankRowIndex)).toBe("");
  });

  it("throws from RowNamed.value and returns empty from RowNamed.valueOrEmpty", () => {
    const row = fetchedDatesSheet().row(blankRowIndex);

    expect(() => row.value("id")).toThrowError(/is empty/);
    expect(row.valueOrEmpty("id")).toBe("");
  });

  it("keeps blanks in the bag returned by RowNamed.valuesOrEmpty", () => {
    const row = fetchedDatesSheet().row(blankRowIndex);

    expect(row.valuesOrEmpty("id", "requiredDate")).toEqual({
      id: "",
      requiredDate: "",
    });
  });

  it("throws from valueArr when a fetched cell is blank, but not from the marked forms", () => {
    const column = fetchedDatesSheet().column("id");

    expect(() => column.valueArr).toThrowError(/is empty/);
    expect(column.valueArrOrEmpty).toEqual(["r:dat:row4", ""]);
    expect(column.valueArrFilterEmpty).toEqual(["r:dat:row4"]);
  });

  it("throws from CellNamed.valueNotEmpty on a blank cell, naming it the same way", () => {
    const cell = fetchedDatesSheet().column("id").cell(blankRowIndex);

    expect(() => cell.valueNotEmpty()).toThrowError(
      new RegExp(`"id".*"dates".*${blankRowIndex}`),
    );
  });

  it("throws from both blank-excluding reads on a column whose box is unticked", () => {
    const sheet = fetchedDatesSheet();
    const column = sheet.column("requiredDate");

    expect(() => column.value(blankRowIndex)).toThrowError(
      new RegExp(`"requiredDate".*"dates".*${blankRowIndex}`),
    );
    expect(() => column.valueNotEmpty(blankRowIndex)).toThrowError(/is empty/);
    expect(() => column.valueArr).toThrowError(/is empty/);
    expect(() => column.valueArrNotEmpty).toThrowError(/is empty/);
    expect(() =>
      sheet.row(blankRowIndex).valueNotEmpty("requiredDate"),
    ).toThrowError(/is empty/);
    expect(column.valueOrEmpty(blankRowIndex)).toBe("");
  });

  it("reads a filled cell identically through all three words", () => {
    const column = fetchedDatesSheet().column("requiredDate");

    expect(column.value(filledRowIndex)).toBe(requiredDateSerial);
    expect(column.valueNotEmpty(filledRowIndex)).toBe(requiredDateSerial);
    expect(column.valueOrEmpty(filledRowIndex)).toBe(requiredDateSerial);
  });

  it("keeps the blank out of the unmarked read's type on a column whose box is unticked", () => {
    const column = fetchedDatesSheet().column("requiredDate");
    // eslint-disable-next-line @typescript-eslint/no-unused-vars -- read for its type
    const cell = column.cell(filledRowIndex);
    // eslint-disable-next-line @typescript-eslint/no-unused-vars -- read for its type
    const row = fetchedDatesSheet().row(filledRowIndex);

    assertType<IsExactly<ReturnType<typeof cell.value>, SerialDate>>(true);
    assertType<IsExactly<ReturnType<typeof cell.valueNotEmpty>, SerialDate>>(
      true,
    );
    assertType<IsExactly<ReturnType<typeof column.value>, SerialDate>>(true);
    assertType<IsExactly<ReturnType<typeof column.valueNotEmpty>, SerialDate>>(
      true,
    );
    assertType<
      IsExactly<ReturnType<typeof column.valueOrEmpty>, SerialDate | "">
    >(true);
    assertType<
      IsExactly<ReturnType<typeof row.value<"requiredDate">>, SerialDate>
    >(true);
  });

  // A plain number would pass an assignment check against SerialDate's supertype.
  it("gives a date column a value type no plain number can be handed to", () => {
    // eslint-disable-next-line @typescript-eslint/no-unused-vars -- read for its type
    const column = fetchedDatesSheet().column("requiredDate");
    // eslint-disable-next-line @typescript-eslint/no-unused-vars -- read for its type
    const numberColumn = fetchedValueTypesSheet().column("numberValue");

    assertNotType<IsExactly<ReturnType<typeof column.value>, number>>(false);
    assertNotType<IsExactly<SerialDate, number>>(false);
    assertType<IsExactly<ReturnType<typeof numberColumn.value>, number>>(true);
  });

  it("hands back the blank rather than throwing on a column whose box is ticked", () => {
    const sheet = fetchedDatesSheet();
    const column = sheet.column("optionalDate");

    expect(column.value(blankRowIndex)).toBe("");
    expect(sheet.row(blankRowIndex).value("optionalDate")).toBe("");
    expect(column.valueArr).toEqual([optionalDateSerial, ""]);
  });

  it("still throws from the blank-excluding reads on a column whose box is ticked", () => {
    const sheet = fetchedDatesSheet();
    const column = sheet.column("optionalDate");

    expect(() => column.valueNotEmpty(blankRowIndex)).toThrowError(/is empty/);
    expect(() => column.valueArrNotEmpty).toThrowError(/is empty/);
    expect(() =>
      sheet.row(blankRowIndex).valueNotEmpty("optionalDate"),
    ).toThrowError(/is empty/);
  });

  it("keeps the blank in the unmarked read's type on a column whose box is ticked", () => {
    const column = fetchedDatesSheet().column("optionalDate");
    // eslint-disable-next-line @typescript-eslint/no-unused-vars -- read for its type
    const cell = column.cell(blankRowIndex);
    // eslint-disable-next-line @typescript-eslint/no-unused-vars -- read for its type
    const row = fetchedDatesSheet().row(blankRowIndex);

    assertType<IsExactly<ReturnType<typeof cell.value>, SerialDate | "">>(true);
    assertType<IsExactly<ReturnType<typeof cell.valueNotEmpty>, SerialDate>>(
      true,
    );
    assertType<IsExactly<ReturnType<typeof column.value>, SerialDate | "">>(
      true,
    );
    assertType<IsExactly<typeof column.valueArr, (SerialDate | "")[]>>(true);
    assertType<IsExactly<ReturnType<typeof column.valueNotEmpty>, SerialDate>>(
      true,
    );
    assertType<
      IsExactly<ReturnType<typeof row.value<"optionalDate">>, SerialDate | "">
    >(true);
  });

  it("lets rowsFiltered select the rows whose column is blank", () => {
    const sheet = fetchedDatesSheet();

    expect(sheet.rowsFiltered({ id: "" }).map((row) => row.rowIndex)).toEqual([
      blankRowIndex,
    ]);
  });

  it("sorts rows by a column that some of them leave blank", () => {
    const sheet = fetchedDatesSheet();

    const sorted = sheet.sortRowsbyColumnName(sheet.rows, "id");

    expect(sorted.map((row) => row.rowIndex)).toEqual([
      blankRowIndex,
      filledRowIndex,
    ]);
  });
});

function fetchedSheetConfig(): SpreadsheetNamed {
  const ss = SpreadsheetNamed.init();
  ss.sheet("sheetConfig").prepFetchColumnsFull(
    "sheetGid",
    "sheetTitle",
    "letApiAccess",
  );
  ss.fetchAllPrepped();
  return ss;
}

function deleteRequestIndexes(
  calls: GoogleAppsScript.Sheets.Schema.BatchUpdateSpreadsheetRequest[],
): (number | undefined)[] {
  return calls
    .flatMap((call) => call.requests ?? [])
    .filter((request) => request.deleteDimension)
    .map((request) => request.deleteDimension?.range?.startIndex);
}

function appendRequestCount(
  calls: GoogleAppsScript.Sheets.Schema.BatchUpdateSpreadsheetRequest[],
): number {
  return calls
    .flatMap((call) => call.requests ?? [])
    .filter((request) => request.appendCells).length;
}

describe("SheetNamed.rowByValue", () => {
  it("returns the one row whose column holds the value", () => {
    stubDatesAndValueTypesWithBlankRow();

    const row = fetchedDatesSheet().rowByValue("id", "r:dat:row4");

    expect(row.rowIndex).toBe(filledRowIndex);
  });

  it("throws naming the sheet, the column and the value when nothing matches", () => {
    stubDatesAndValueTypesWithBlankRow();
    const sheet = fetchedDatesSheet();

    expect(() => sheet.rowByValue("id", "r:dat:absent")).toThrowError(
      /dates.*id.*r:dat:absent.*0 did/,
    );
  });

  it("throws rather than picking one when two rows match", () => {
    stubDatesWithDuplicateIds();
    const sheet = fetchedDatesSheet();

    expect(() => sheet.rowByValue("id", "r:dat:dup")).toThrowError(/but 2 did/);
  });

  // The blank row reads "" through valueOrEmpty, so it is a match like any other.
  it("counts the sheet's blank row as a match for an empty value", () => {
    stubDatesAndValueTypesWithBlankRow();

    const row = fetchedDatesSheet().rowByValue("id", "");

    expect(row.rowIndex).toBe(blankRowIndex);
  });
});

function stubDatesWithDuplicateIds() {
  return stubSheetsService({
    sheets: [
      {
        sheetId: datesGid,
        title: "Dates",
        rows: buildGridRows({
          0: [idColumnId, requiredDateColumnId, optionalDateColumnId],
          3: ["ID", "Required date", "Optional date"],
          4: ["r:dat:dup", requiredDateSerial, optionalDateSerial],
          5: ["r:dat:dup", requiredDateSerial, optionalDateSerial],
        }),
        table: { endRowIndex: 6 },
      },
    ],
  });
}

describe("SheetNamed.DELETE_ALL_DATA_ROWS", () => {
  it("deletes every data row but the top one, and leaves that one blank", () => {
    const { batchUpdateCalls } = stubSheetConfigSheet({
      4: filledSheetConfigRow,
      5: filledSheetConfigRow,
      6: filledSheetConfigRow,
    });

    const ss = fetchedSheetConfig();
    ss.sheet("sheetConfig").DELETE_ALL_DATA_ROWS();
    ss.batchUpdateGSheets();

    expect(deleteRequestIndexes(batchUpdateCalls)).toEqual([6, 5]);
    expect(ss.sheet("sheetConfig").topRow.isBlank).toBe(true);
  });

  it("writes nothing at all for a sheet already down to its blank row", () => {
    const { batchUpdateCalls } = stubSheetConfigSheet({
      4: blankSheetConfigRow,
    });

    const ss = fetchedSheetConfig();
    ss.sheet("sheetConfig").DELETE_ALL_DATA_ROWS();
    ss.batchUpdateGSheets();

    expect(batchUpdateCalls).toEqual([]);
  });
});

describe("SheetNamed.appendRowWithVals", () => {
  it("reuses the blank row of an emptied sheet rather than appending beneath it", () => {
    const { batchUpdateCalls } = stubSheetConfigSheet({
      4: blankSheetConfigRow,
    });

    const ss = fetchedSheetConfig();
    const row = ss.sheet("sheetConfig").appendRowWithVals({
      sheetTitle: "Item",
    });
    ss.batchUpdateGSheets();

    expect(row.rowIndex).toBe(topDataRowIndex);
    expect(appendRequestCount(batchUpdateCalls)).toBe(0);
    expect(row.value("sheetTitle")).toBe("Item");
  });

  it("appends beneath a one-row sheet that still holds data", () => {
    const { batchUpdateCalls } = stubSheetConfigSheet({
      4: filledSheetConfigRow,
    });

    const ss = fetchedSheetConfig();
    const row = ss
      .sheet("sheetConfig")
      .appendRowWithVals({ sheetTitle: "Log" });
    ss.batchUpdateGSheets();

    expect(row.rowIndex).toBe(topDataRowIndex + 1);
    expect(appendRequestCount(batchUpdateCalls)).toBe(1);
  });

  it("reuses the blank row once and appends for the second row", () => {
    const { batchUpdateCalls } = stubSheetConfigSheet({
      4: blankSheetConfigRow,
    });

    const ss = fetchedSheetConfig();
    const sheet = ss.sheet("sheetConfig");
    const first = sheet.appendRowWithVals({ sheetTitle: "one" });
    const second = sheet.appendRowWithVals({ sheetTitle: "two" });
    ss.batchUpdateGSheets();

    expect([first.rowIndex, second.rowIndex]).toEqual([
      topDataRowIndex,
      topDataRowIndex + 1,
    ]);
    expect(appendRequestCount(batchUpdateCalls)).toBe(1);
  });

  // The wipe has to lift the reservation the first append took, or the second strands a row.
  it("hands the same row to a second append once a wipe has released it", () => {
    const { batchUpdateCalls } = stubSheetConfigSheet({
      4: blankSheetConfigRow,
    });

    const ss = fetchedSheetConfig();
    const sheet = ss.sheet("sheetConfig");
    sheet.appendRowWithVals({ sheetTitle: "one" });
    sheet.DELETE_ALL_DATA_ROWS();
    const rebuilt = sheet.appendRowWithVals({ sheetTitle: "two" });
    ss.batchUpdateGSheets();

    expect(rebuilt.rowIndex).toBe(topDataRowIndex);
    expect(appendRequestCount(batchUpdateCalls)).toBe(0);
    expect(rebuilt.value("sheetTitle")).toBe("two");
  });

  it("reuses the row a wipe just cleared, so the wipe and rebuild leave only rebuilt rows", () => {
    const { batchUpdateCalls } = stubSheetConfigSheet({
      4: filledSheetConfigRow,
      5: filledSheetConfigRow,
    });

    const ss = fetchedSheetConfig();
    const sheet = ss.sheet("sheetConfig");
    sheet.DELETE_ALL_DATA_ROWS();
    const row = sheet.appendRowWithVals({ sheetTitle: "new" });
    ss.batchUpdateGSheets();

    expect(row.rowIndex).toBe(topDataRowIndex);
    expect(appendRequestCount(batchUpdateCalls)).toBe(0);
    expect(deleteRequestIndexes(batchUpdateCalls)).toEqual([5]);
  });
});

const valueTypesColumnIdRow = Object.values({
  id: getColumnTraitByName("valueTypes", "id", "columnId"),
  stringValue: getColumnTraitByName("valueTypes", "stringValue", "columnId"),
  numberValue: getColumnTraitByName("valueTypes", "numberValue", "columnId"),
  dateValue: getColumnTraitByName("valueTypes", "dateValue", "columnId"),
  sampledBoolean: getColumnTraitByName(
    "valueTypes",
    "sampledBoolean",
    "columnId",
  ),
  checkbox: getColumnTraitByName("valueTypes", "checkbox", "columnId"),
});

function stubValueTypesWithBlankRow() {
  return stubSheetsService({
    sheets: [
      {
        sheetId: valueTypesGid,
        title: "Value Types",
        rows: buildGridRows({
          0: valueTypesColumnIdRow,
          4: valueTypesColumnIdRow.map(() => null),
        }),
        table: { endRowIndex: 5 },
      },
    ],
  });
}

function fetchedValueTypesSpreadsheet(): SpreadsheetNamed {
  const ss = SpreadsheetNamed.init();
  ss.sheet("valueTypes").prepFetchColumnsFull(
    "id",
    "stringValue",
    "numberValue",
    "dateValue",
    "sampledBoolean",
    "checkbox",
  );
  ss.fetchAllPrepped();
  return ss;
}

type CompleteAppendBag<SN extends SheetName> = Parameters<
  SheetNamed<SN>["appendRowWithAllVals"]
>[0];

const completeValueTypesRow: CompleteAppendBag<"valueTypes"> = {
  stringValue: "Yes",
  numberValue: 7,
  dateValue: SerialDate.fromYmd({ year: 2024, month: 3, day: 14 }),
  sampledBoolean: true,
  checkbox: true,
};

describe("SheetNamed.appendRowWithAllVals", () => {
  it("mints the row ID itself, from a bag that cannot name one", () => {
    stubValueTypesWithBlankRow();

    const row = fetchedValueTypesSpreadsheet()
      .sheet("valueTypes")
      .appendRowWithAllVals(completeValueTypesRow);

    expect(row.value("id")).toMatch(/^r:vty:[0-9a-zA-Z_-]{7}$/);
  });

  it("writes every value the bag carries", () => {
    stubValueTypesWithBlankRow();

    const row = fetchedValueTypesSpreadsheet()
      .sheet("valueTypes")
      .appendRowWithAllVals(completeValueTypesRow);

    expect([
      row.value("stringValue"),
      row.value("numberValue"),
      row.value("dateValue"),
      row.value("sampledBoolean"),
      row.value("checkbox"),
    ]).toEqual([
      "Yes",
      7,
      SerialDate.fromYmd({ year: 2024, month: 3, day: 14 }),
      true,
      true,
    ]);
  });

  it("reuses the blank row the way the partial append does", () => {
    const { batchUpdateCalls } = stubValueTypesWithBlankRow();

    const ss = fetchedValueTypesSpreadsheet();
    const row = ss
      .sheet("valueTypes")
      .appendRowWithAllVals(completeValueTypesRow);
    ss.batchUpdateGSheets();

    expect(row.rowIndex).toBe(topDataRowIndex);
    expect(appendRequestCount(batchUpdateCalls)).toBe(0);
  });

  it("asks a sheet with an ID column for every writable column but the ID", () => {
    assertType<
      IsExactly<
        keyof CompleteAppendBag<"valueTypes">,
        | "stringValue"
        | "numberValue"
        | "dateValue"
        | "sampledBoolean"
        | "checkbox"
      >
    >(true);
    assertType<
      IsExactly<CompleteAppendBag<"valueTypes">["numberValue"], number | "">
    >(true);
  });

  it("asks a sheet with no ID column for every writable column", () => {
    assertType<
      IsExactly<
        CompleteAppendBag<"sheetConfig">,
        {
          sheetGid: number | "";
          sheetTitle: string;
          letApiAccess: boolean;
        }
      >
    >(true);
  });

  it("refuses a bag that names the ID", () => {
    const withId: CompleteAppendBag<"valueTypes"> = {
      ...completeValueTypesRow,
      // @ts-expect-error the append mints the ID, so a caller cannot supply one
      id: "r:vty:abcdefg",
    };

    expect(Object.keys(withId)).toEqual([
      "stringValue",
      "numberValue",
      "dateValue",
      "sampledBoolean",
      "checkbox",
      "id",
    ]);
  });
});

// Both dates are blank: the required one is a gap, the optional one is allowed.
function fetchedBlankDatesRow(): RowNamed<"dates"> {
  stubSheetsService({
    sheets: [
      {
        sheetId: datesGid,
        title: "Dates",
        rows: buildGridRows({
          0: [idColumnId, requiredDateColumnId, optionalDateColumnId],
          3: ["ID", "Required date", "Optional date"],
          4: ["r:dat:row4", null, null],
        }),
        table: { endRowIndex: 5 },
      },
    ],
  });
  const ss = SpreadsheetNamed.init();
  ss.sheet("dates").prepFetchColumnsFull("id", "requiredDate", "optionalDate");
  ss.fetchAllPrepped();
  return ss.sheet("dates").row(topDataRowIndex);
}

describe("RowNamed.blankRequiredColumnNames", () => {
  it("names a blank column whose Empty value allowed box is unticked", () => {
    expect(fetchedBlankDatesRow().blankRequiredColumnNames()).toContain(
      "requiredDate",
    );
  });

  it("leaves out a blank column whose box is ticked", () => {
    expect(fetchedBlankDatesRow().blankRequiredColumnNames()).not.toContain(
      "optionalDate",
    );
  });

  it("leaves out an unticked checkbox, whose blank is an answer", () => {
    stubDatesAndValueTypesWithBlankRow();
    const row = fetchedValueTypesSheet().row(blankRowIndex);

    expect(row.blankRequiredColumnNames()).not.toContain("checkbox");
    expect(row.blankRequiredColumnNames()).toContain("numberValue");
  });

  it("names every blank required column and nothing else", () => {
    expect(fetchedBlankDatesRow().blankRequiredColumnNames()).toEqual([
      "requiredDate",
    ]);
  });
});

const computedGid = getSheetTraitByName("computed", "sheetGid");
const computedColumnIdRow = [
  getColumnTraitByName("computed", "amount", "columnId"),
  getColumnTraitByName("computed", "rowNumber", "columnId"),
];
const testFormula = "=2+SINGLE(computed[Amount])";
const rowNumberColIndex = 1;

function stubComputedForFormulaWrite() {
  return stubSheetsService({
    sheets: [
      {
        sheetId: computedGid,
        title: "Computed",
        rows: buildGridRows({
          0: computedColumnIdRow,
          4: [10, 11],
          5: [20, 21],
        }),
        table: { endRowIndex: 6 },
      },
    ],
  });
}

describe("Named formula writes", () => {
  it("sends one pasteData PASTE_FORMULA for every Computed data row", () => {
    const { batchUpdateCalls } = stubComputedForFormulaWrite();

    const ss = SpreadsheetNamed.init();
    ss.fetchAllSheetProperties();
    ss.sheet("computed").column("rowNumber").updateAllFormulas(testFormula);
    ss.batchUpdateGSheets();

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        pasteData: {
          coordinate: {
            sheetId: computedGid,
            rowIndex: topDataRowIndex,
            columnIndex: rowNumberColIndex,
          },
          data: `"${testFormula}"\n"${testFormula}"`,
          delimiter: "\t",
          type: "PASTE_FORMULA",
        },
      },
    ]);
  });

  it("sends one pasteData PASTE_FORMULA for a single Row number cell", () => {
    const { batchUpdateCalls } = stubComputedForFormulaWrite();

    const ss = SpreadsheetNamed.init();
    ss.fetchAllSheetProperties();
    ss.sheet("computed")
      .column("rowNumber")
      .cell(topDataRowIndex)
      .updateFormula(testFormula);
    ss.batchUpdateGSheets();

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        pasteData: {
          coordinate: {
            sheetId: computedGid,
            rowIndex: topDataRowIndex,
            columnIndex: rowNumberColIndex,
          },
          data: `"${testFormula}"`,
          delimiter: "\t",
          type: "PASTE_FORMULA",
        },
      },
    ]);
  });

  it("sends one pasteData per contiguous active run for updateActiveFormulas", () => {
    const { batchUpdateCalls } = stubComputedForFormulaWrite();

    const ss = SpreadsheetNamed.init();
    ss.sheet("computed").prepFetchColumnsFull("rowNumber");
    ss.fetchAllPrepped();
    ss.sheet("computed").raw.removeRowsExcept(topDataRowIndex);
    ss.sheet("computed").column("rowNumber").updateActiveFormulas(testFormula);
    ss.batchUpdateGSheets();

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        pasteData: {
          coordinate: {
            sheetId: computedGid,
            rowIndex: topDataRowIndex,
            columnIndex: rowNumberColIndex,
          },
          data: `"${testFormula}"`,
          delimiter: "\t",
          type: "PASTE_FORMULA",
        },
      },
    ]);
  });

  it("throws before queueing when the formula does not start with =", () => {
    stubComputedForFormulaWrite();

    const ss = SpreadsheetNamed.init();
    ss.fetchAllSheetProperties();

    expect(() =>
      ss.sheet("computed").column("rowNumber").updateAllFormulas("2+1"),
    ).toThrowError('Formula must start with "=". Got "2+1".');
  });

  it("leaves local cell values unchanged after a formula write", () => {
    stubComputedForFormulaWrite();

    const ss = SpreadsheetNamed.init();
    ss.sheet("computed").prepFetchColumnsFull("rowNumber", "amount");
    ss.fetchAllPrepped();
    ss.sheet("computed").column("rowNumber").updateAllFormulas(testFormula);

    expect(ss.sheet("computed").column("rowNumber").valueArrOrEmpty).toEqual([
      11, 21,
    ]);
    expect(ss.sheet("computed").column("amount").valueArrOrEmpty).toEqual([
      10, 20,
    ]);
  });

  it("still refuses a value write on Row number", () => {
    stubComputedForFormulaWrite();

    const ss = SpreadsheetNamed.init();
    ss.fetchAllSheetProperties();

    expect(() =>
      ss
        .sheet("computed")
        .column("rowNumber")
        .cell(topDataRowIndex)
        .updateValue(99),
    ).toThrowError(/formula column/);
  });

  it("merges a colour onto the same cell as a formula write", () => {
    const { batchUpdateCalls } = stubComputedForFormulaWrite();
    const backgroundColor = { red: 0.851, green: 0.918, blue: 0.827 };

    const ss = SpreadsheetNamed.init();
    ss.fetchAllSheetProperties();
    const cell = ss.sheet("computed").column("rowNumber").cell(topDataRowIndex);
    cell.updateFormula(testFormula);
    cell.updateBackgroundColor(backgroundColor);
    ss.batchUpdateGSheets();

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        pasteData: {
          coordinate: {
            sheetId: computedGid,
            rowIndex: topDataRowIndex,
            columnIndex: rowNumberColIndex,
          },
          data: `"${testFormula}"`,
          delimiter: "\t",
          type: "PASTE_FORMULA",
        },
      },
      {
        updateCells: {
          range: {
            sheetId: computedGid,
            startRowIndex: topDataRowIndex,
            endRowIndex: topDataRowIndex + 1,
            startColumnIndex: rowNumberColIndex,
            endColumnIndex: rowNumberColIndex + 1,
          },
          rows: [
            {
              values: [{ userEnteredFormat: { backgroundColor } }],
            },
          ],
          fields: "userEnteredFormat.backgroundColor",
        },
      },
    ]);
  });

  it("refuses a whole-column formula fill on a sheet pruned to a selection", () => {
    stubComputedForFormulaWrite();

    const ss = SpreadsheetNamed.init();
    ss.sheet("computed").prepFetchColumnsFull("rowNumber");
    ss.fetchAllPrepped();
    ss.sheet("computed").raw.removeRowsExcept(topDataRowIndex);

    expect(() =>
      ss.sheet("computed").column("rowNumber").updateAllFormulas(testFormula),
    ).toThrowError(/pruned to a selection/);
  });

  it("accepts Row number and rejects Amount at the type level", () => {
    assertType<IsExactly<ColumnIsFormula<"computed", "rowNumber">, true>>(true);
    assertType<IsExactly<ColumnIsFormula<"computed", "amount">, false>>(true);

    function formulaWriteTypeGate(
      formulaColumn: ColumnNamed<"computed", "rowNumber">,
      numColumn: ColumnNamed<"computed", "amount">,
    ) {
      formulaColumn.updateAllFormulas(testFormula);
      formulaColumn.updateActiveFormulas(testFormula);
      formulaColumn.cell(topDataRowIndex).updateFormula(testFormula);
      // @ts-expect-error Amount is not a formula column
      numColumn.updateAllFormulas(testFormula);
      // @ts-expect-error Amount is not a formula column
      numColumn.updateActiveFormulas(testFormula);
      // @ts-expect-error Amount is not a formula column
      numColumn.cell(topDataRowIndex).updateFormula(testFormula);
    }

    expect(formulaWriteTypeGate).toEqual(expect.any(Function));
  });

  it("rejects a misspelled column type at the type level", () => {
    function columnTypeGate(column: ColumnNamed<"computed", "amount">) {
      column.meta.updateColumnType("DOUBLE");
      // @ts-expect-error DOUBEL is not a Table column type
      column.meta.updateColumnType("DOUBEL");
    }

    expect(columnTypeGate).toEqual(expect.any(Function));
  });
});
