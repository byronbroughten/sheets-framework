import { describe, expect, it } from "vitest";

import { googleRawRequest } from "../00_Source/GoogleSheets/GoogleSheetsAPI";
import { installedRawSource } from "../00_Source/RawSource/RawSource";
import { expectedSheetLayout } from "./expectedSheetLayout";
import {
  buildGridRows,
  type FakeSheetProperties,
  type FakeTable,
  stubSheetsService,
} from "./fakeSheetsService";
import { a1Notation } from "./fakeSheetsService/a1Notation";
import type { BoundedRange } from "./fakeSheetsService/fakeGrid";
import type { FakeTablePlacement } from "./fakeSheetsService/fakeTables";
import { fakeTableSheet } from "./fakeSheetsService/fakeTableSheet";
import type { FakeGridView } from "./fakeSheetsService/gridView";

type Request = GoogleAppsScript.Sheets.Schema.Request;

const widgetGid = 7;

// Mirrors the dev-spreadsheet check in docs/testing.md: a Table at A1:C4, a marker below it.
function widgetSheet(): FakeSheetProperties {
  return {
    sheetId: widgetGid,
    title: "Widget",
    rows: buildGridRows({
      0: ["h1", "h2", "h3"],
      1: [1, "x", "p"],
      2: [2, "y", "q"],
      3: [3, "z", "r"],
      6: ["below"],
    }),
    table: {
      name: "Widget",
      startRowIndex: 0,
      startColumnIndex: 0,
      endRowIndex: 4,
      endColumnIndex: 3,
      columnTypes: { 0: "DOUBLE", 1: "TEXT" },
    },
  };
}

function send(...requests: Request[]): void {
  installedRawSource().flush(
    requests.map((request) => ({
      kind: "raw",
      request: googleRawRequest(request),
    })),
  );
}

function values(
  ...cells: (string | number)[]
): GoogleAppsScript.Sheets.Schema.RowData {
  return {
    values: cells.map((cell) => ({
      userEnteredValue:
        typeof cell === "number"
          ? { numberValue: cell }
          : { stringValue: cell },
    })),
  };
}

function rowRange(startIndex: number, endIndex: number) {
  return { sheetId: widgetGid, dimension: "ROWS", startIndex, endIndex };
}

describe("stubSheetsService replays appends where the live API places them", () => {
  it("inserts a Table append at the Table's end, pushing the rows below down", () => {
    const { grid } = stubSheetsService({ sheets: [widgetSheet()] });

    send({
      appendCells: {
        sheetId: widgetGid,
        tableId: "fake-table-7",
        rows: [values(4, "w")],
        fields: "userEnteredValue",
      },
    });

    const sheet = grid.sheet(widgetGid);
    expect(sheet.values({ startRowIndex: 4, endRowIndex: 8 })).toEqual([
      [4, "w", null],
      [null, null, null],
      [null, null, null],
      ["below", null, null],
    ]);
    expect(sheet.tables[0]?.range?.endRowIndex).toBe(5);
    expect(sheet.rowCount).toBe(8);
  });

  it("puts a sheet append after the last row holding data", () => {
    const { grid } = stubSheetsService({ sheets: [widgetSheet()] });

    send({
      appendCells: {
        sheetId: widgetGid,
        rows: [values("tail")],
        fields: "userEnteredValue",
      },
    });

    expect(grid.sheet(widgetGid).values({ startRowIndex: 6 })).toEqual([
      ["below", null, null],
      ["tail", null, null],
    ]);
  });
});

describe("stubSheetsService replays row and column changes", () => {
  it("grows a Table by an insert inside it but not by one at its end", () => {
    const { grid } = stubSheetsService({ sheets: [widgetSheet()] });

    send(
      { insertDimension: { range: rowRange(4, 5), inheritFromBefore: false } },
      { insertDimension: { range: rowRange(2, 3), inheritFromBefore: false } },
    );

    const sheet = grid.sheet(widgetGid);
    expect(sheet.tables[0]?.range?.endRowIndex).toBe(5);
    expect(sheet.values({ startColumnIndex: 0, endColumnIndex: 1 })).toEqual([
      ["h1"],
      [1],
      [null],
      [2],
      [3],
      [null],
      [null],
      [null],
      ["below"],
    ]);
  });

  it("grows a Table by a column inserted at its end only when it inherits from before, and heads it Column <n>", () => {
    const { grid } = stubSheetsService({ sheets: [widgetSheet()] });

    send({
      insertDimension: {
        range: {
          sheetId: widgetGid,
          dimension: "COLUMNS",
          startIndex: 3,
          endIndex: 4,
        },
        inheritFromBefore: true,
      },
    });

    const table = grid.sheet(widgetGid).tables[0];
    expect(table?.range?.endColumnIndex).toBe(4);
    expect(table?.columnProperties?.map((column) => column.columnName)).toEqual(
      ["h1", "h2", "h3", "Column 4"],
    );
  });

  it("copies each left neighbour's format and validation, but not its value or column type, into a column inserted inheriting from before", () => {
    const green = { red: 0.2, green: 0.8, blue: 0.2 };
    const { grid } = stubSheetsService({
      sheets: [
        {
          ...widgetSheet(),
          rows: buildGridRows({
            0: ["h1", "h2", "h3"],
            1: [
              1,
              "x",
              { value: true, dataValidationConditionType: "BOOLEAN" },
            ],
            2: [2, "y", { value: 3, numberFormatType: "CURRENCY" }],
            3: [3, "z", { value: "r", backgroundColor: green }],
          }),
          table: {
            name: "Widget",
            startRowIndex: 0,
            startColumnIndex: 0,
            endRowIndex: 4,
            endColumnIndex: 3,
            columnTypes: { 2: "BOOLEAN" },
          },
        },
      ],
    });

    send({
      insertDimension: {
        range: {
          sheetId: widgetGid,
          dimension: "COLUMNS",
          startIndex: 3,
          endIndex: 4,
        },
        inheritFromBefore: true,
      },
    });

    const sheet = grid.sheet(widgetGid);
    expect(sheet.rows({ startRowIndex: 1, startColumnIndex: 3 })).toEqual([
      [{ value: false, dataValidationConditionType: "BOOLEAN" }],
      [{ value: null, numberFormatType: "CURRENCY" }],
      [{ value: null, backgroundColor: green }],
    ]);
    expect(sheet.tables[0]?.columnProperties?.[3]?.columnType).toBeUndefined();
  });

  it("moves a column's type with the column an insert shifts", () => {
    const { grid } = stubSheetsService({ sheets: [widgetSheet()] });

    send({
      insertDimension: {
        range: {
          sheetId: widgetGid,
          dimension: "COLUMNS",
          startIndex: 1,
          endIndex: 2,
        },
      },
    });

    const columns = grid.sheet(widgetGid).tables[0]?.columnProperties;
    expect(columns?.map((column) => column.columnType)).toEqual([
      "DOUBLE",
      undefined,
      "TEXT",
      undefined,
    ]);
  });

  it("drops a deleted row and shrinks the Table around it", () => {
    const { grid } = stubSheetsService({ sheets: [widgetSheet()] });

    send({ deleteDimension: { range: rowRange(1, 2) } });

    const sheet = grid.sheet(widgetGid);
    expect(sheet.values({ endRowIndex: 3 })).toEqual([
      ["h1", "h2", "h3"],
      [2, "y", "q"],
      [3, "z", "r"],
    ]);
    expect(sheet.tables[0]?.range?.endRowIndex).toBe(3);
  });

  it("hides and unhides rows", () => {
    const { grid } = stubSheetsService({ sheets: [widgetSheet()] });

    send(
      {
        updateDimensionProperties: {
          range: rowRange(1, 4),
          properties: { hiddenByUser: true },
          fields: "hiddenByUser",
        },
      },
      {
        updateDimensionProperties: {
          range: rowRange(2, 3),
          properties: { hiddenByUser: false },
          fields: "hiddenByUser",
        },
      },
    );

    expect(grid.sheet(widgetGid).hiddenRowIndexes).toEqual([1, 3]);
  });

  it("sorts descending as the reverse of numbers, text, booleans, with blanks still last", () => {
    const { grid } = stubSheetsService({
      sheets: [
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: [["text"], [true], [], [3], [false]],
        },
      ],
    });

    send({
      sortRange: {
        range: { sheetId: widgetGid, startRowIndex: 0, endRowIndex: 5 },
        sortSpecs: [{ dimensionIndex: 0, sortOrder: "DESCENDING" }],
      },
    });

    expect(grid.sheet(widgetGid).values({ endRowIndex: 5 })).toEqual([
      [true],
      [false],
      ["text"],
      [3],
      [null],
    ]);
  });

  it("sorts numbers before text and blanks last", () => {
    const { grid } = stubSheetsService({
      sheets: [
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: [["text", "a"], [], [10, "b"], [3, "c"]],
        },
      ],
    });

    send({
      sortRange: {
        range: { sheetId: widgetGid, startRowIndex: 0, endRowIndex: 4 },
        sortSpecs: [{ dimensionIndex: 0, sortOrder: "ASCENDING" }],
      },
    });

    expect(grid.sheet(widgetGid).values({ endRowIndex: 4 })).toEqual([
      [3, "c"],
      [10, "b"],
      ["text", "a"],
      [null, null],
    ]);
  });
});

describe("stubSheetsService replays Table updates by their field mask", () => {
  it("renames a Table without touching its columns", () => {
    const { grid } = stubSheetsService({ sheets: [widgetSheet()] });

    send({
      updateTable: {
        table: { tableId: "fake-table-7", name: "Renamed" },
        fields: "name",
      },
    });

    const table = grid.sheet(widgetGid).tables[0];
    expect(table?.name).toBe("Renamed");
    expect(table?.columnProperties?.[0]?.columnType).toBe("DOUBLE");
  });

  it("keeps an unsent column's header but drops its type, as a partial column list does live", () => {
    const { grid } = stubSheetsService({ sheets: [widgetSheet()] });

    send({
      updateTable: {
        table: {
          tableId: "fake-table-7",
          columnProperties: [{ columnIndex: 1, columnName: "h2b" }],
        },
        fields: "columnProperties",
      },
    });

    expect(grid.sheet(widgetGid).tables[0]?.columnProperties).toEqual([
      { columnName: "h1" },
      { columnIndex: 1, columnName: "h2b" },
      { columnIndex: 2, columnName: "h3" },
    ]);
  });

  // Measured live with room below the Table: the rows it takes in stay blank, formula columns included.
  it("extends a Table's range over blank rows without filling them, formula columns included", () => {
    const { grid } = stubSheetsService({
      sheets: [
        {
          ...widgetSheet(),
          rows: buildGridRows({
            0: ["h1", "h2", "h3"],
            1: [1, "x", { value: 2, isFormula: true, formula: "=A2+1" }],
            2: [2, "y", { value: 3, isFormula: true, formula: "=A3+1" }],
            3: [3, "z", { value: 4, isFormula: true, formula: "=A4+1" }],
            7: ["below"],
          }),
        },
      ],
    });

    send({
      updateTable: {
        table: {
          tableId: "fake-table-7",
          range: {
            sheetId: widgetGid,
            startRowIndex: 0,
            endRowIndex: 6,
            startColumnIndex: 0,
            endColumnIndex: 3,
          },
        },
        fields: "range",
      },
    });

    expect(grid.sheet(widgetGid).tables[0]?.range).toMatchObject({
      startRowIndex: 0,
      endRowIndex: 6,
    });
    expect(
      grid.sheet(widgetGid).values({ startRowIndex: 4, endRowIndex: 6 }),
    ).toEqual([
      [null, null, null],
      [null, null, null],
    ]);
  });

  it("throws naming a validation-rule field beyond the condition's type and values", () => {
    stubSheetsService({ sheets: [widgetSheet()] });

    expect(() =>
      send({
        updateTable: {
          table: {
            tableId: "fake-table-7",
            columnProperties: [
              {
                columnName: "h1",
                dataValidationRule: {
                  condition: { type: "BOOLEAN" },
                  strict: true,
                },
              },
            ],
          },
          fields: "columnProperties",
        },
      }),
    ).toThrowError('dataValidationRule field "strict"');
  });

  it("adds a second Table beside the first", () => {
    const { grid } = stubSheetsService({ sheets: [widgetSheet()] });

    send({
      addTable: {
        table: {
          tableId: "side",
          name: "Side",
          range: {
            sheetId: widgetGid,
            startRowIndex: 0,
            endRowIndex: 2,
            startColumnIndex: 4,
            endColumnIndex: 5,
          },
        },
      },
    });

    expect(grid.sheet(widgetGid).tables.map((table) => table.name)).toEqual([
      "Widget",
      "Side",
    ]);
  });
});

describe("stubSheetsService replays deleteTable as the live API does", () => {
  const deletedGid = 9;
  const tableRange = {
    sheetId: deletedGid,
    startRowIndex: 3,
    endRowIndex: 6,
    startColumnIndex: 0,
    endColumnIndex: 2,
  };

  function stubSheetWithHeadedTable() {
    return stubSheetsService({
      sheets: [
        {
          sheetId: deletedGid,
          title: "Deleted",
          rows: buildGridRows({
            0: ["c:lse:aaa", "c:lse:bbb"],
            1: ["Heading"],
            3: ["ID", "Status"],
            4: [
              { value: 1, numberFormatType: "CURRENCY" },
              { value: "x", dataValidationConditionType: "BOOLEAN" },
            ],
            5: [{ value: 2, isFormula: true, formula: "=1+1" }, "y"],
            6: ["below"],
          }),
          table: {
            name: "Deleted",
            ...tableRange,
            columnTypes: { 0: "DOUBLE" },
          },
          conditionalFormats: [
            { ranges: [{ ...tableRange, startRowIndex: 4 }] },
            { ranges: [{ ...tableRange, endRowIndex: 7 }] },
            {
              ranges: [
                { sheetId: deletedGid, startColumnIndex: 0, endColumnIndex: 1 },
              ],
            },
          ],
        },
      ],
    });
  }

  it("drops the Table and blanks its whole range, leaving the cells above and below", () => {
    const { grid } = stubSheetWithHeadedTable();

    send({ deleteTable: { tableId: "fake-table-9" } });

    const sheet = grid.sheet(deletedGid);
    expect(sheet.tables).toEqual([]);
    expect(sheet.rows({ startRowIndex: 0, endRowIndex: 7 })).toEqual([
      ["c:lse:aaa", "c:lse:bbb"],
      ["Heading", null],
      [null, null],
      [null, null],
      [null, null],
      [null, null],
      ["below", null],
    ]);
  });

  it("rewrites an outside formula's Table reference to the A1 range of the column's old body rows, never #REF!", () => {
    const { grid } = stubSheetsService({
      sheets: [
        {
          sheetId: deletedGid,
          title: "Deleted",
          rows: buildGridRows({
            3: ["ID", "Status"],
            4: [1, "x"],
            5: [2, "y"],
            0: [
              null,
              null,
              null,
              { value: "=SUM(Deleted[ID])", isFormula: true },
              { value: "=ROWS(Deleted[Status])", isFormula: true },
            ],
          }),
          table: { name: "Deleted", ...tableRange },
        },
      ],
    });

    send({ deleteTable: { tableId: "fake-table-9" } });

    expect(
      grid.sheet(deletedGid).values({
        startRowIndex: 0,
        endRowIndex: 1,
        startColumnIndex: 3,
        endColumnIndex: 5,
      }),
    ).toEqual([["=SUM(Deleted!$A$5:$A$6)", "=ROWS(Deleted!$B$5:$B$6)"]]);
  });

  it("leaves another Table's reference alone when its name only ends with the deleted Table's, and quotes a title with a dollar sign literally", () => {
    const { grid } = stubSheetsService({
      sheets: [
        {
          sheetId: deletedGid,
          title: "Pay $& Tax",
          rows: buildGridRows({
            3: ["ID", "Status"],
            4: [1, "x"],
            0: [
              null,
              null,
              null,
              { value: "=SUM(Deleted[ID])", isFormula: true },
              { value: "=SUM(myDeleted[ID])", isFormula: true },
            ],
          }),
          table: { name: "Deleted", ...tableRange },
        },
      ],
    });

    send({ deleteTable: { tableId: "fake-table-9" } });

    expect(
      grid.sheet(deletedGid).values({
        startRowIndex: 0,
        endRowIndex: 1,
        startColumnIndex: 3,
        endColumnIndex: 5,
      }),
    ).toEqual([["=SUM('Pay $& Tax'!$A$5:$A$6)", "=SUM(myDeleted[ID])"]]);
  });

  it("throws on an outside reference in a form it doesn't rewrite, rather than leaving it as written", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: deletedGid,
          title: "Deleted",
          rows: buildGridRows({
            3: ["ID", "Status"],
            4: [1, "x"],
            0: [
              null,
              null,
              null,
              { value: "=Deleted[[#ALL]]", isFormula: true },
            ],
          }),
          table: { name: "Deleted", ...tableRange },
        },
      ],
    });

    expect(() =>
      send({ deleteTable: { tableId: "fake-table-9" } }),
    ).toThrowError('reading Table "Deleted" in a form other than name[header]');
  });

  it("removes the rules lying wholly over the Table and keeps those reaching past it", () => {
    const { grid } = stubSheetWithHeadedTable();

    send({ deleteTable: { tableId: "fake-table-9" } });

    expect(
      grid.sheet(deletedGid).conditionalFormats.map((rule) => rule.ranges),
    ).toEqual([
      [{ ...tableRange, endRowIndex: 7 }],
      [{ sheetId: deletedGid, startColumnIndex: 0, endColumnIndex: 1 }],
    ]);
  });
});

const ruleGid = 111;

describe("stubSheetsService replays updateConditionalFormatRule", () => {
  function stubRules() {
    return stubSheetsService({
      sheets: [
        {
          sheetId: ruleGid,
          title: "Rules",
          conditionalFormats: [
            { ranges: [{ sheetId: ruleGid, startColumnIndex: 0 }] },
            { ranges: [{ sheetId: ruleGid, startColumnIndex: 1 }] },
          ],
        },
      ],
    });
  }

  it("replaces the rule at the index and leaves the others", () => {
    const { grid } = stubRules();

    send({
      updateConditionalFormatRule: {
        sheetId: ruleGid,
        index: 1,
        rule: { ranges: [{ sheetId: ruleGid, startColumnIndex: 5 }] },
      },
    });

    expect(
      grid.sheet(ruleGid).conditionalFormats.map((rule) => rule.ranges),
    ).toEqual([
      [{ sheetId: ruleGid, startColumnIndex: 0 }],
      [{ sheetId: ruleGid, startColumnIndex: 5 }],
    ]);
  });

  it("throws on an index past the sheet's last rule", () => {
    stubRules();

    expect(() =>
      send({
        updateConditionalFormatRule: {
          sheetId: ruleGid,
          index: 2,
          rule: { ranges: [] },
        },
      }),
    ).toThrow(
      `updateConditionalFormatRule: sheet ${ruleGid} has no rule at index 2.`,
    );
  });
});

describe("stubSheetsService replays tabs", () => {
  it("adds, renames and deletes a tab, and a later read sees it", () => {
    const { grid } = stubSheetsService({ sheets: [widgetSheet()] });

    send(
      {
        addSheet: {
          properties: {
            sheetId: 8,
            title: "Gadget",
            gridProperties: { rowCount: 5, columnCount: 2 },
          },
        },
      },
      {
        updateSheetProperties: {
          properties: { sheetId: widgetGid, title: "Widgets" },
          fields: "title",
        },
      },
    );
    const titlesRead = installedRawSource()
      .fetchSheetProperties()
      .sheets.map((sheet) => sheet.title);
    send({ deleteSheet: { sheetId: widgetGid } });

    expect(titlesRead).toEqual(["Widgets", "Gadget"]);
    expect(grid.sheetTitles()).toEqual(["Gadget"]);
    expect(grid.sheet(8).rowCount).toBe(5);
  });
});

describe("stubSheetsService replays cell writes", () => {
  it("writes only the fields a mask names", () => {
    const { grid } = stubSheetsService({
      sheets: [
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: [[{ value: 45000, numberFormatType: "DATE" }, "keep"]],
        },
      ],
    });
    const red = { red: 1, green: 0, blue: 0 };

    send(
      {
        updateCells: {
          range: {
            sheetId: widgetGid,
            startRowIndex: 0,
            endRowIndex: 1,
            startColumnIndex: 0,
            endColumnIndex: 1,
          },
          rows: [values(46000)],
          fields: "userEnteredValue",
        },
      },
      {
        repeatCell: {
          range: {
            sheetId: widgetGid,
            startRowIndex: 0,
            endRowIndex: 2,
            startColumnIndex: 1,
            endColumnIndex: 2,
          },
          cell: { userEnteredFormat: { backgroundColor: red } },
          fields: "userEnteredFormat.backgroundColor",
        },
      },
      {
        setDataValidation: {
          range: {
            sheetId: widgetGid,
            startRowIndex: 1,
            endRowIndex: 2,
            startColumnIndex: 0,
            endColumnIndex: 1,
          },
          rule: { condition: { type: "BOOLEAN" } },
        },
      },
    );

    expect(grid.sheet(widgetGid).rows()).toEqual([
      [
        { value: 46000, numberFormatType: "DATE" },
        { value: "keep", backgroundColor: red },
      ],
      [
        { value: null, dataValidationConditionType: "BOOLEAN" },
        { value: null, backgroundColor: red },
      ],
    ]);
  });

  it("pastes each record into its own row, a formula as its text", () => {
    const { grid } = stubSheetsService({ sheets: [widgetSheet()] });

    send({
      pasteData: {
        coordinate: { sheetId: widgetGid, rowIndex: 1, columnIndex: 2 },
        data: '"=A2+1"\n"=A3+1"',
        delimiter: "\t",
        type: "PASTE_FORMULA",
      },
    });

    expect(
      grid.sheet(widgetGid).rows({
        startRowIndex: 1,
        endRowIndex: 3,
        startColumnIndex: 2,
        endColumnIndex: 3,
      }),
    ).toEqual([
      [{ value: "=A2+1", isFormula: true }],
      [{ value: "=A3+1", isFormula: true }],
    ]);
  });

  it("replaces text across every sheet, leaving numbers alone", () => {
    const { grid } = stubSheetsService({ sheets: [widgetSheet()] });

    send({
      findReplace: {
        find: "^h(\\d)$",
        replacement: "header $1",
        searchByRegex: true,
        allSheets: true,
      },
    });

    expect(grid.sheet(widgetGid).values({ endRowIndex: 2 })).toEqual([
      ["header 1", "header 2", "header 3"],
      [1, "x", "p"],
    ]);
  });
});

const bandGid = 9;
const partOfTable = "You cannot insert or delete cells over part of a table.";
const modelGreen = { red: 0, green: 1, blue: 0 };

interface BandTables {
  side?: FakeTable;
  lower?: FakeTable;
}

// Mirrors the live probes (sheets-framework#53, #56, #59): Table A at A4:C6 with its head rows, markers below and beside it.
function bandSheet({ side, lower }: BandTables = {}): FakeSheetProperties {
  return {
    sheetId: bandGid,
    title: "Band",
    rows: buildGridRows({
      4: [1, 2, 3],
      5: [
        { value: 4, backgroundColor: modelGreen, numberFormatType: "CURRENCY" },
        { value: 5, dataValidationConditionType: "BOOLEAN" },
        6,
      ],
      6: ["belowA", null, null, null, "beside"],
    }),
    tables: [
      {
        tableId: "a",
        name: "A",
        startRowIndex: 3,
        startColumnIndex: 0,
        endRowIndex: 6,
        endColumnIndex: 3,
        headRows: { 3: ["a:1", "a:2", "a:3"], 0: ["One", "Two", "Three"] },
      },
      ...(side === undefined ? [] : [side]),
      ...(lower === undefined ? [] : [lower]),
    ],
  };
}

function sideTable(endRowIndex: number): FakeTable {
  return {
    tableId: "b",
    name: "B",
    startRowIndex: 3,
    startColumnIndex: 5,
    endRowIndex,
    endColumnIndex: 7,
    headRows: { 0: ["Left", "Right"] },
  };
}

function lowerTable(endColumnIndex: number): FakeTable {
  return {
    tableId: "c",
    name: "C",
    startRowIndex: 10,
    startColumnIndex: 0,
    endRowIndex: 12,
    endColumnIndex,
    headRows: { 3: ["c:1"], 0: ["Low"] },
  };
}

function bandRange(
  range: BoundedRange,
): GoogleAppsScript.Sheets.Schema.GridRange {
  return { sheetId: bandGid, ...range };
}

function tableRange(
  grid: FakeGridView,
  tableId: string,
): GoogleAppsScript.Sheets.Schema.GridRange | undefined {
  return grid.sheet(bandGid).tables.find((table) => table.tableId === tableId)
    ?.range;
}

describe("stubSheetsService places each fixture Table with its own head rows", () => {
  it("writes a Table's head rows above its own header, from its first column", () => {
    const { grid } = stubSheetsService({
      sheets: [bandSheet({ lower: lowerTable(1) })],
    });

    const sheet = grid.sheet(bandGid);
    expect(sheet.values({ endRowIndex: 1, endColumnIndex: 3 })).toEqual([
      ["a:1", "a:2", "a:3"],
    ]);
    expect(
      sheet.values({ startRowIndex: 7, endRowIndex: 11, endColumnIndex: 1 }),
    ).toEqual([["c:1"], [null], [null], ["Low"]]);
    expect(sheet.tables.map((table) => table.tableId)).toEqual(["a", "c"]);
  });

  it("names unnamed Tables by their place on the sheet", () => {
    const { grid } = stubSheetsService({
      sheets: [
        {
          sheetId: bandGid,
          title: "Band",
          tables: [{ endRowIndex: 5 }, { startRowIndex: 8, endRowIndex: 9 }],
        },
      ],
    });

    expect(grid.sheet(bandGid).tables.map((table) => table.tableId)).toEqual([
      "fake-table-9",
      "fake-table-9-1",
    ]);
  });

  it("throws on a head-row cell the fixture's rows already hold", () => {
    expect(() =>
      stubSheetsService({
        sheets: [
          {
            ...bandSheet(),
            rows: buildGridRows({ 0: ["clash"] }),
          },
        ],
      }),
    ).toThrowError(/head row 3 and rows both give a cell at row 0, column 0/);
  });
});

describe("stubSheetsService replays Table-bounded inserts and deletes", () => {
  it("grows the grid by appended rows without moving a cell", () => {
    const { grid } = stubSheetsService({ sheets: [bandSheet()] });
    const before = grid.sheet(bandGid).values();

    send({
      appendDimension: { sheetId: bandGid, dimension: "ROWS", length: 3 },
    });

    expect(grid.sheet(bandGid).rowCount).toBe(10);
    expect(grid.sheet(bandGid).values()).toEqual(before);
  });

  it("answers a fetch with the row count appended rows left", () => {
    stubSheetsService({ sheets: [bandSheet()] });

    send({
      appendDimension: { sheetId: bandGid, dimension: "ROWS", length: 3 },
    });

    expect(
      installedRawSource().fetchSheetProperties().sheets[0]?.rowCount,
    ).toBe(10);
  });

  it("pushes down only the cells in an inserted range's columns, and grows no Table the insert sits just below", () => {
    const { grid } = stubSheetsService({ sheets: [bandSheet()] });

    send(
      { appendDimension: { sheetId: bandGid, dimension: "ROWS", length: 2 } },
      {
        insertRange: {
          range: bandRange({
            startRowIndex: 6,
            endRowIndex: 8,
            startColumnIndex: 0,
            endColumnIndex: 3,
          }),
          shiftDimension: "ROWS",
        },
      },
    );

    const sheet = grid.sheet(bandGid);
    expect(sheet.values({ startRowIndex: 6, endRowIndex: 9 })).toEqual([
      [null, null, null, null, "beside"],
      [null, null, null, null, null],
      ["belowA", null, null, null, null],
    ]);
    expect(tableRange(grid, "a")?.endRowIndex).toBe(6);
    expect(sheet.rowCount).toBe(9);
  });

  it("grows a Table by a row range inserted inside it", () => {
    const { grid } = stubSheetsService({ sheets: [bandSheet()] });

    send({
      insertRange: {
        range: bandRange({
          startRowIndex: 5,
          endRowIndex: 6,
          startColumnIndex: 0,
          endColumnIndex: 3,
        }),
        shiftDimension: "ROWS",
      },
    });

    expect(tableRange(grid, "a")?.endRowIndex).toBe(7);
  });

  it("grows the grid as far as the cells an insert pushes past its edge", () => {
    const { grid } = stubSheetsService({ sheets: [bandSheet()] });

    send({
      insertRange: {
        range: bandRange({
          startRowIndex: 4,
          endRowIndex: 6,
          startColumnIndex: 0,
          endColumnIndex: 3,
        }),
        shiftDimension: "ROWS",
      },
    });

    expect(grid.sheet(bandGid).cell(8, 0)).toBe("belowA");
    expect(grid.sheet(bandGid).rowCount).toBe(9);
  });

  it("clips a range that runs past the grid's edge to the grid, as live", () => {
    const { grid } = stubSheetsService({ sheets: [bandSheet()] });

    send({
      insertRange: {
        range: bandRange({
          startRowIndex: 6,
          endRowIndex: 8,
          startColumnIndex: 0,
          endColumnIndex: 3,
        }),
        shiftDimension: "ROWS",
      },
    });

    expect(grid.sheet(bandGid).cell(7, 0)).toBe("belowA");
    expect(grid.sheet(bandGid).rowCount).toBe(8);
  });

  it("refuses, as live, a range that starts past the grid's edge", () => {
    stubSheetsService({ sheets: [bandSheet()] });

    expect(() =>
      send({
        deleteRange: {
          range: bandRange({
            startRowIndex: 7,
            endRowIndex: 9,
            startColumnIndex: 0,
            endColumnIndex: 3,
          }),
          shiftDimension: "ROWS",
        },
      }),
    ).toThrowError(
      "Invalid requests[0].deleteRange: Range ('Band'!A8:C9) exceeds grid limits. Max rows: 7, max columns: 5",
    );
  });

  it("pushes a narrower Table below down intact, head rows included", () => {
    const { grid } = stubSheetsService({
      sheets: [bandSheet({ lower: lowerTable(1) })],
    });

    send({
      insertRange: {
        range: bandRange({
          startRowIndex: 6,
          endRowIndex: 8,
          startColumnIndex: 0,
          endColumnIndex: 3,
        }),
        shiftDimension: "ROWS",
      },
    });

    expect(tableRange(grid, "c")).toEqual({
      startRowIndex: 12,
      endRowIndex: 14,
      startColumnIndex: 0,
      endColumnIndex: 1,
    });
    expect(grid.sheet(bandGid).cell(9, 0)).toBe("c:1");
  });

  it("refuses, as live, a row insert that would split a wider Table below", () => {
    const { grid } = stubSheetsService({
      sheets: [bandSheet({ lower: lowerTable(5) })],
    });

    expect(() =>
      send({
        insertRange: {
          range: bandRange({
            startRowIndex: 6,
            endRowIndex: 7,
            startColumnIndex: 0,
            endColumnIndex: 3,
          }),
          shiftDimension: "ROWS",
        },
      }),
    ).toThrowError(`Invalid requests[0].insertRange: ${partOfTable}`);
    expect(grid.sheet(bandGid).cell(6, 0)).toBe("belowA");
  });

  it("shifts a column range right past a Table without widening it, carrying a side Table no taller", () => {
    const { grid } = stubSheetsService({
      sheets: [bandSheet({ side: sideTable(6) })],
    });

    send({
      insertRange: {
        range: bandRange({
          startRowIndex: 0,
          endRowIndex: 6,
          startColumnIndex: 3,
          endColumnIndex: 4,
        }),
        shiftDimension: "COLUMNS",
      },
    });

    const sheet = grid.sheet(bandGid);
    expect(tableRange(grid, "a")?.endColumnIndex).toBe(3);
    expect(tableRange(grid, "b")).toEqual({
      startRowIndex: 3,
      endRowIndex: 6,
      startColumnIndex: 6,
      endColumnIndex: 8,
    });
    expect(sheet.values({ startRowIndex: 3, endRowIndex: 4 })).toEqual([
      ["One", "Two", "Three", null, null, null, "Left", "Right"],
    ]);
    expect(sheet.cell(6, 4)).toBe("beside");
  });

  it("refuses, as live, a column insert beside a taller Table", () => {
    stubSheetsService({ sheets: [bandSheet({ side: sideTable(9) })] });

    expect(() =>
      send({
        insertRange: {
          range: bandRange({
            startRowIndex: 0,
            endRowIndex: 6,
            startColumnIndex: 3,
            endColumnIndex: 4,
          }),
          shiftDimension: "COLUMNS",
        },
      }),
    ).toThrowError(`Invalid requests[0].insertRange: ${partOfTable}`);
  });

  it("shrinks a Table by a row range deleted over its columns, sparing the Table beside it", () => {
    const { grid } = stubSheetsService({
      sheets: [bandSheet({ side: sideTable(6) })],
    });

    send({
      deleteRange: {
        range: bandRange({
          startRowIndex: 4,
          endRowIndex: 5,
          startColumnIndex: 0,
          endColumnIndex: 3,
        }),
        shiftDimension: "ROWS",
      },
    });

    const sheet = grid.sheet(bandGid);
    expect(tableRange(grid, "a")?.endRowIndex).toBe(5);
    expect(tableRange(grid, "b")?.endRowIndex).toBe(6);
    expect(
      sheet.values({ startRowIndex: 4, endRowIndex: 7, endColumnIndex: 5 }),
    ).toEqual([
      [4, 5, 6, null, null],
      ["belowA", null, null, null, null],
      [null, null, null, null, "beside"],
    ]);
    expect(sheet.rowCount).toBe(7);
  });

  it("keeps a whole-sheet protection whole under a row delete over some columns", () => {
    const wholeSheet = { protectedRangeId: 7, range: { sheetId: bandGid } };
    const { grid } = stubSheetsService({
      sheets: [{ ...bandSheet(), protectedRanges: [wholeSheet] }],
    });

    send({
      deleteRange: {
        range: bandRange({
          startRowIndex: 4,
          endRowIndex: 5,
          startColumnIndex: 0,
          endColumnIndex: 3,
        }),
        shiftDimension: "ROWS",
      },
    });

    expect(grid.sheet(bandGid).protectedRanges).toEqual([wholeSheet]);
  });

  it("leaves a Table header only when a delete takes its every body row", () => {
    const { grid } = stubSheetsService({ sheets: [bandSheet()] });

    send({
      deleteRange: {
        range: bandRange({
          startRowIndex: 4,
          endRowIndex: 6,
          startColumnIndex: 0,
          endColumnIndex: 3,
        }),
        shiftDimension: "ROWS",
      },
    });

    expect(tableRange(grid, "a")).toEqual({
      startRowIndex: 3,
      endRowIndex: 4,
      startColumnIndex: 0,
      endColumnIndex: 3,
    });
  });

  it("refuses, as live, a delete of a Table's header row", () => {
    stubSheetsService({ sheets: [bandSheet()] });

    expect(() =>
      send({
        deleteRange: {
          range: bandRange({
            startRowIndex: 3,
            endRowIndex: 4,
            startColumnIndex: 0,
            endColumnIndex: 3,
          }),
          shiftDimension: "ROWS",
        },
      }),
    ).toThrowError(
      "Invalid requests[0].deleteRange: Cannot delete a table header row. Consider hiding the row instead.",
    );
  });

  it("refuses, as live, a delete over part of a Table's columns", () => {
    stubSheetsService({ sheets: [bandSheet({ side: sideTable(6) })] });

    expect(() =>
      send({
        deleteRange: {
          range: bandRange({
            startRowIndex: 4,
            endRowIndex: 5,
            startColumnIndex: 0,
            endColumnIndex: 6,
          }),
          shiftDimension: "ROWS",
        },
      }),
    ).toThrowError(`Invalid requests[0].deleteRange: ${partOfTable}`);
  });
});

const measuredGid = 217;

// Mirrors the live probes (sheets-framework#217, Scratch217A and E to G, K): Table at A4:E7 under four rules.
function measuredRuleSheet(): FakeSheetProperties {
  return {
    sheetId: measuredGid,
    title: "Rules217",
    rows: buildGridRows({ 6: [1, 2, 3, 4, 5] }),
    tables: [
      {
        tableId: "m",
        name: "M",
        startRowIndex: 3,
        startColumnIndex: 0,
        endRowIndex: 7,
        endColumnIndex: 5,
        headRows: { 0: ["One", "Two", "Three", "Four", "Five"] },
      },
    ],
    conditionalFormats: [
      formulaRule("A5:E7", '=$C5="x"'),
      formulaRule("D5:D7", "=D5>0"),
      formulaRule("B4:B7", "=LEN(B4)>5"),
      formulaRule("E5:E7", "=D5>2"),
    ],
  };
}

function measuredRange(a1: string): GoogleAppsScript.Sheets.Schema.GridRange {
  return a1GridRange(measuredGid, a1);
}

function a1GridRange(
  sheetId: number,
  a1: string,
): GoogleAppsScript.Sheets.Schema.GridRange {
  const match = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(a1);
  const [, startCol = "A", startRow = "1", endCol = "A", endRow = "1"] =
    match ?? [];
  return {
    sheetId,
    startRowIndex: Number(startRow) - 1,
    endRowIndex: Number(endRow),
    startColumnIndex: a1Notation.columnIndex(startCol),
    endColumnIndex: a1Notation.columnIndex(endCol) + 1,
  };
}

function formulaRule(
  a1: string,
  formula: string,
): GoogleAppsScript.Sheets.Schema.ConditionalFormatRule {
  return {
    ranges: [measuredRange(a1)],
    booleanRule: {
      condition: {
        type: "CUSTOM_FORMULA",
        values: [{ userEnteredValue: formula }],
      },
      format: {},
    },
  };
}

function measuredBandColumns(
  startColumnIndex: number,
): GoogleAppsScript.Sheets.Schema.GridRange {
  return {
    sheetId: measuredGid,
    startRowIndex: 0,
    endRowIndex: 7,
    startColumnIndex,
    endColumnIndex: startColumnIndex + 1,
  };
}

const cutUnderRuleMessage =
  "The fake Sheets service does not replay a cutPaste under a conditional-format rule, other than a sideways cut inside a whole-Table rule naming no cut or pasted column.";

function measuredRules(
  grid: FakeGridView,
): GoogleAppsScript.Sheets.Schema.ConditionalFormatRule[] {
  return grid.sheet(measuredGid).conditionalFormats;
}

describe("stubSheetsService replays rules under Table-bounded column inserts and deletes as measured", () => {
  it("grows a whole-row rule, moves the rules right of the insert with their references, and leaves an anchored reference as written", () => {
    const { grid } = stubSheetsService({ sheets: [measuredRuleSheet()] });

    send({
      insertRange: {
        range: measuredBandColumns(2),
        shiftDimension: "COLUMNS",
      },
    });

    expect(measuredRules(grid)).toEqual([
      formulaRule("A5:F7", '=$C5="x"'),
      formulaRule("E5:E7", "=E5>0"),
      formulaRule("B4:B7", "=LEN(B4)>5"),
      formulaRule("F5:F7", "=E5>2"),
    ]);
  });

  it("removes a rule over only the deleted column, keeps a whole-row rule's range, and leaves a reference to the deleted column as written", () => {
    const { grid } = stubSheetsService({ sheets: [measuredRuleSheet()] });

    send({
      deleteRange: {
        range: measuredBandColumns(3),
        shiftDimension: "COLUMNS",
      },
    });

    expect(measuredRules(grid)).toEqual([
      formulaRule("A5:E7", '=$C5="x"'),
      formulaRule("B4:B7", "=LEN(B4)>5"),
      formulaRule("D5:D7", "=D5>2"),
    ]);
  });

  it("moves a relative reference right of a delete left, and leaves an anchored reference to the deleted column as written", () => {
    const { grid } = stubSheetsService({ sheets: [measuredRuleSheet()] });

    send({
      deleteRange: {
        range: measuredBandColumns(2),
        shiftDimension: "COLUMNS",
      },
    });

    expect(measuredRules(grid)).toEqual([
      formulaRule("A5:E7", '=$C5="x"'),
      formulaRule("C5:C7", "=C5>0"),
      formulaRule("B4:B7", "=LEN(B4)>5"),
      formulaRule("D5:D7", "=C5>2"),
    ]);
  });

  it("throws naming a rule spanning a column change without spanning a whole Table", () => {
    stubSheetsService({
      sheets: [
        {
          ...measuredRuleSheet(),
          conditionalFormats: [formulaRule("B5:D7", "=$B5>0")],
        },
      ],
    });

    expect(() =>
      send({
        insertRange: {
          range: measuredBandColumns(2),
          shiftDimension: "COLUMNS",
        },
      }),
    ).toThrowError(
      "The fake Sheets service does not replay a column change inside a conditional-format rule over part of a Table (B5:D7).",
    );
  });

  it("throws naming an anchored reference in a rule a column change moves", () => {
    stubSheetsService({
      sheets: [
        {
          ...measuredRuleSheet(),
          conditionalFormats: [formulaRule("D5:D7", "=$D5>0")],
        },
      ],
    });

    expect(() =>
      send({
        insertRange: {
          range: measuredBandColumns(2),
          shiftDimension: "COLUMNS",
        },
      }),
    ).toThrowError(
      "The fake Sheets service does not replay a column change under the reference $D5 in a conditional-format rule over D5:D7.",
    );
  });

  it.each([
    ["=COUNTIF(D:D,D5)>1", "D:D"],
    ["=COUNTIF(5:5,D5)>1", "5:5"],
    ["=Other!D5>1", "!"],
  ])(
    "throws naming an unmeasured reference shape in %s",
    (formula, reference) => {
      stubSheetsService({
        sheets: [
          {
            ...measuredRuleSheet(),
            conditionalFormats: [formulaRule("D5:D7", formula)],
          },
        ],
      });

      expect(() =>
        send({
          insertRange: {
            range: measuredBandColumns(2),
            shiftDimension: "COLUMNS",
          },
        }),
      ).toThrowError(
        `The fake Sheets service does not replay a column change under the reference ${reference} in a conditional-format rule over D5:D7.`,
      );
    },
  );

  it("throws naming a rule part of whose columns a delete takes", () => {
    stubSheetsService({
      sheets: [
        {
          ...measuredRuleSheet(),
          conditionalFormats: [formulaRule("D5:E7", "=D5>0")],
        },
      ],
    });

    expect(() =>
      send({
        deleteRange: {
          range: measuredBandColumns(3),
          shiftDimension: "COLUMNS",
        },
      }),
    ).toThrowError(
      "The fake Sheets service does not replay a column delete over part of a conditional-format rule (D5:E7).",
    );
  });
});

describe("stubSheetsService replays rules under a Table-bounded column move as measured", () => {
  // Scratch217A: the first column moved to before the last, by an insert, a cut into the gap and a delete.
  const moveFirstColumnBeforeLast: Request[] = [
    {
      insertRange: { range: measuredBandColumns(4), shiftDimension: "COLUMNS" },
    },
    {
      cutPaste: {
        source: measuredBandColumns(0),
        destination: { sheetId: measuredGid, rowIndex: 0, columnIndex: 4 },
        pasteType: "PASTE_NORMAL",
      },
    },
    {
      deleteRange: { range: measuredBandColumns(0), shiftDimension: "COLUMNS" },
    },
  ];

  it("moves column rules and their relative references, and leaves a whole-row rule one column past the Table, its anchored reference as written", () => {
    const { grid } = stubSheetsService({ sheets: [measuredRuleSheet()] });

    send(...moveFirstColumnBeforeLast);

    expect(measuredRules(grid)).toEqual([
      formulaRule("A5:F7", '=$C5="x"'),
      formulaRule("C5:C7", "=C5>0"),
      formulaRule("A4:A7", "=LEN(A4)>5"),
      formulaRule("E5:E7", "=C5>2"),
    ]);
  });

  it.each([['=$A5="x"'], ['=A5="x"'], ['=$E5="x"']])(
    "throws on a whole-row rule %s naming a cut or pasted column",
    (formula) => {
      stubSheetsService({
        sheets: [
          {
            ...measuredRuleSheet(),
            conditionalFormats: [formulaRule("A5:E7", formula)],
          },
        ],
      });

      expect(() => send(...moveFirstColumnBeforeLast)).toThrowError(
        cutUnderRuleMessage,
      );
    },
  );

  it("leaves all four rules matching the columns' new places once the whole-row rule is updated at its index (Scratch217K)", () => {
    const { grid } = stubSheetsService({ sheets: [measuredRuleSheet()] });

    send(...moveFirstColumnBeforeLast, {
      updateConditionalFormatRule: {
        sheetId: measuredGid,
        index: 0,
        rule: formulaRule("A5:E7", '=$B5="x"'),
      },
    });

    expect(measuredRules(grid)).toEqual([
      formulaRule("A5:E7", '=$B5="x"'),
      formulaRule("C5:C7", "=C5>0"),
      formulaRule("A4:A7", "=LEN(A4)>5"),
      formulaRule("E5:E7", "=C5>2"),
    ]);
  });
});

const edgeGid = 337;

interface EdgeTable {
  tableId: string;
  startColumnIndex?: number;
  endRowIndex?: number;
  headers: string[];
  columnIds?: string[];
}

function edgeTable({
  tableId,
  startColumnIndex = 0,
  endRowIndex = 6,
  headers,
  columnIds,
}: EdgeTable): FakeTable {
  return {
    tableId,
    name: tableId,
    startRowIndex: 3,
    startColumnIndex,
    endRowIndex,
    endColumnIndex: startColumnIndex + headers.length,
    headRows: {
      0: headers,
      ...(columnIds === undefined ? {} : { 3: columnIds }),
    },
  };
}

function edgeRange(a1: string): GoogleAppsScript.Sheets.Schema.GridRange {
  return a1GridRange(edgeGid, a1);
}

function edgeTableRanges(
  grid: FakeGridView,
): GoogleAppsScript.Sheets.Schema.GridRange[] {
  return grid
    .sheet(edgeGid)
    .tables.map((table) => ({ sheetId: edgeGid, ...table.range }));
}

describe("stubSheetsService replays the grid-edge measurements (sheets-framework#337)", () => {
  it("grows a Table on the grid's last row by every appended row, blank, and leaves no row outside it (Scratch337A)", () => {
    const { grid } = stubSheetsService({
      sheets: [
        {
          sheetId: edgeGid,
          title: "Scratch337A",
          rows: buildGridRows({ 4: ["a", 1], 5: ["b", 2] }),
          tables: [
            edgeTable({ tableId: "s337a", headers: ["Name", "Amount"] }),
          ],
        },
      ],
    });

    send({
      appendDimension: { sheetId: edgeGid, dimension: "ROWS", length: 2 },
    });

    const sheet = grid.sheet(edgeGid);
    expect(sheet.rowCount).toBe(8);
    expect(edgeTableRanges(grid)).toEqual([edgeRange("A4:B8")]);
    expect(
      sheet.values({ startRowIndex: 6, endRowIndex: 8, endColumnIndex: 2 }),
    ).toEqual([
      [null, null],
      [null, null],
    ]);
  });

  it("grows a side-by-side neighbour ending on the grid's last row by a blank row when only one Table is appended to (Scratch337B)", () => {
    const { grid } = stubSheetsService({
      sheets: [
        {
          sheetId: edgeGid,
          title: "Scratch337B",
          rows: buildGridRows({
            4: ["a", 1, null, "x", 10],
            5: ["b", 2, null, "y", 20],
          }),
          tables: [
            edgeTable({ tableId: "s337l", headers: ["Name", "Amount"] }),
            edgeTable({
              tableId: "s337r",
              startColumnIndex: 3,
              headers: ["Label", "Count"],
            }),
          ],
        },
      ],
    });

    send(
      { appendDimension: { sheetId: edgeGid, dimension: "ROWS", length: 1 } },
      {
        updateTable: {
          table: { tableId: "s337l", range: edgeRange("A4:B7") },
          fields: "range",
        },
      },
      {
        updateCells: {
          range: edgeRange("A7:B7"),
          rows: [values("c", 3)],
          fields: "userEnteredValue",
        },
      },
    );

    const sheet = grid.sheet(edgeGid);
    expect(sheet.rowCount).toBe(7);
    expect(edgeTableRanges(grid)).toEqual([
      edgeRange("A4:B7"),
      edgeRange("D4:E7"),
    ]);
    expect(sheet.values({ startRowIndex: 6 })).toEqual([
      ["c", 3, null, null, null],
    ]);
  });

  it("keeps the cells a move's insert pushes past the grid's last column, leaving the grid a blank column wider (Scratch337C)", () => {
    const { grid } = stubSheetsService({
      sheets: [
        {
          sheetId: edgeGid,
          title: "Scratch337C",
          rows: buildGridRows({
            4: ["a1", "b1", "c1", "d1"],
            5: ["a2", "b2", "c2", "d2"],
            6: ["a3", "b3", "c3", "d3"],
          }),
          tables: [
            edgeTable({
              tableId: "s337c",
              endRowIndex: 7,
              headers: ["One", "Two", "Three", "Four"],
              columnIds: ["c:1", "c:2", "c:3", "c:4"],
            }),
          ],
        },
      ],
    });

    send(
      {
        insertRange: { range: edgeRange("B1:B7"), shiftDimension: "COLUMNS" },
      },
      {
        cutPaste: {
          source: edgeRange("E1:E7"),
          destination: { sheetId: edgeGid, rowIndex: 0, columnIndex: 1 },
          pasteType: "PASTE_NORMAL",
        },
      },
      {
        deleteRange: { range: edgeRange("E1:E7"), shiftDimension: "COLUMNS" },
      },
    );

    const sheet = grid.sheet(edgeGid);
    expect(sheet.columnCount).toBe(5);
    expect(edgeTableRanges(grid)).toEqual([edgeRange("A4:D7")]);
    expect(sheet.values({ endRowIndex: 1, endColumnIndex: 5 })).toEqual([
      ["c:1", "c:4", "c:2", "c:3", null],
    ]);
    expect(sheet.values({ startRowIndex: 4, endColumnIndex: 5 })).toEqual([
      ["a1", "d1", "b1", "c1", null],
      ["a2", "d2", "b2", "c2", null],
      ["a3", "d3", "b3", "c3", null],
    ]);
  });
});

describe("stubSheetsService replays copyPaste of format and validation", () => {
  const modelRow = bandRange({
    startRowIndex: 5,
    endRowIndex: 6,
    startColumnIndex: 0,
    endColumnIndex: 3,
  });
  const newRows = bandRange({
    startRowIndex: 6,
    endRowIndex: 8,
    startColumnIndex: 0,
    endColumnIndex: 3,
  });

  it("tiles the model row's format over the new rows, leaving their values and validation", () => {
    const { grid } = stubSheetsService({
      sheets: [
        {
          ...bandSheet(),
          conditionalFormats: [
            {
              ranges: [
                bandRange({
                  startRowIndex: 4,
                  endRowIndex: 6,
                  startColumnIndex: 0,
                  endColumnIndex: 1,
                }),
              ],
            },
            {
              ranges: [
                { sheetId: bandGid, startColumnIndex: 1, endColumnIndex: 2 },
              ],
            },
          ],
        },
      ],
    });

    send(
      { appendDimension: { sheetId: bandGid, dimension: "ROWS", length: 1 } },
      {
        copyPaste: {
          source: modelRow,
          destination: newRows,
          pasteType: "PASTE_FORMAT",
          pasteOrientation: "NORMAL",
        },
      },
    );

    const sheet = grid.sheet(bandGid);
    expect(
      sheet.rows({ startRowIndex: 6, endRowIndex: 8, endColumnIndex: 2 }),
    ).toEqual([
      [
        {
          value: "belowA",
          backgroundColor: modelGreen,
          numberFormatType: "CURRENCY",
        },
        null,
      ],
      [
        {
          value: null,
          backgroundColor: modelGreen,
          numberFormatType: "CURRENCY",
        },
        null,
      ],
    ]);
    expect(sheet.conditionalFormats.map((rule) => rule.ranges)).toEqual([
      [
        bandRange({
          startRowIndex: 4,
          endRowIndex: 8,
          startColumnIndex: 0,
          endColumnIndex: 1,
        }),
      ],
      [{ sheetId: bandGid, startColumnIndex: 1, endColumnIndex: 2 }],
    ]);
  });

  it("tiles the model row's cell validation over the new rows, and nothing else", () => {
    const { grid } = stubSheetsService({ sheets: [bandSheet()] });

    send(
      { appendDimension: { sheetId: bandGid, dimension: "ROWS", length: 1 } },
      {
        copyPaste: {
          source: modelRow,
          destination: newRows,
          pasteType: "PASTE_DATA_VALIDATION",
          pasteOrientation: "NORMAL",
        },
      },
    );

    expect(
      grid
        .sheet(bandGid)
        .rows({ startRowIndex: 6, endRowIndex: 8, endColumnIndex: 2 }),
    ).toEqual([
      ["belowA", { value: null, dataValidationConditionType: "BOOLEAN" }],
      [null, { value: null, dataValidationConditionType: "BOOLEAN" }],
    ]);
  });

  it("clips a copy's destination to the grid, as live", () => {
    const { grid } = stubSheetsService({ sheets: [bandSheet()] });

    send({
      copyPaste: {
        source: modelRow,
        destination: newRows,
        pasteType: "PASTE_FORMAT",
        pasteOrientation: "NORMAL",
      },
    });

    expect(grid.sheet(bandGid).rowCount).toBe(7);
  });

  it("throws naming a paste type it does not replay", () => {
    stubSheetsService({ sheets: [bandSheet()] });

    expect(() =>
      send({
        copyPaste: {
          source: modelRow,
          destination: newRows,
          pasteType: "PASTE_NORMAL",
        },
      }),
    ).toThrowError(
      "The fake Sheets service does not replay copyPaste type PASTE_NORMAL.",
    );
  });
});

describe("stubSheetsService replays cutPaste as a whole-cell move", () => {
  const cutBody = bandRange({
    startRowIndex: 4,
    endRowIndex: 6,
    startColumnIndex: 0,
    endColumnIndex: 3,
  });
  const intoColumn3 = { sheetId: bandGid, rowIndex: 4, columnIndex: 3 };

  function cut(
    overrides: Partial<GoogleAppsScript.Sheets.Schema.CutPasteRequest> = {},
  ): Request {
    return {
      cutPaste: {
        source: cutBody,
        destination: intoColumn3,
        pasteType: "PASTE_NORMAL",
        ...overrides,
      },
    };
  }

  it("moves each cell whole, value, format and validation, and leaves the source empty", () => {
    const { grid } = stubSheetsService({ sheets: [bandSheet()] });

    send(cut());

    expect(
      grid
        .sheet(bandGid)
        .rows({ startRowIndex: 4, endRowIndex: 6, endColumnIndex: 6 }),
    ).toEqual([
      [null, null, null, 1, 2, 3],
      [
        null,
        null,
        null,
        { value: 4, backgroundColor: modelGreen, numberFormatType: "CURRENCY" },
        { value: 5, dataValidationConditionType: "BOOLEAN" },
        6,
      ],
    ]);
  });

  it("carries a Table's column types and validation when the cut holds its whole rows", () => {
    const sheet = bandSheet();
    const { grid } = stubSheetsService({
      sheets: [
        {
          ...sheet,
          tables: (sheet.tables ?? []).map((table) => ({
            ...table,
            endColumnIndex: 6,
            columnTypes: { 0: "BOOLEAN", 1: "DROPDOWN" },
            columnValidationConditionTypes: { 1: "ONE_OF_LIST" },
            columnValidationValues: { 1: ["x", "y"] },
          })),
        },
      ],
    });

    send(
      cut({
        source: bandRange({
          startRowIndex: 0,
          endRowIndex: 6,
          startColumnIndex: 0,
          endColumnIndex: 2,
        }),
        destination: { sheetId: bandGid, rowIndex: 0, columnIndex: 3 },
      }),
    );

    const columns = grid.sheet(bandGid).tables[0]?.columnProperties ?? [];
    expect(columns.map((column) => column.columnType)).toEqual([
      undefined,
      undefined,
      undefined,
      "BOOLEAN",
      "DROPDOWN",
      undefined,
    ]);
    expect(columns[4]?.dataValidationRule?.condition).toEqual({
      type: "ONE_OF_LIST",
      values: [{ userEnteredValue: "x" }, { userEnteredValue: "y" }],
    });
    expect(columns[1]?.dataValidationRule).toBeUndefined();
  });

  it("throws naming a paste type it does not replay", () => {
    stubSheetsService({ sheets: [bandSheet()] });

    expect(() => send(cut({ pasteType: "PASTE_FORMAT" }))).toThrowError(
      "The fake Sheets service does not replay cutPaste type PASTE_FORMAT.",
    );
  });

  it("throws on a cut across sheets", () => {
    stubSheetsService({ sheets: [bandSheet()] });

    expect(() =>
      send(cut({ destination: { ...intoColumn3, sheetId: bandGid + 1 } })),
    ).toThrowError(
      "The fake Sheets service does not replay cutPaste across sheets.",
    );
  });

  it("throws on a destination overlapping its own source", () => {
    stubSheetsService({ sheets: [bandSheet()] });

    expect(() =>
      send(cut({ destination: { ...intoColumn3, columnIndex: 2 } })),
    ).toThrowError(
      "The fake Sheets service does not replay a cutPaste onto its own source.",
    );
  });

  it("throws on a source under a conditional-format rule", () => {
    stubSheetsService({
      sheets: [
        {
          ...bandSheet(),
          conditionalFormats: [{ ranges: [cutBody] }],
        },
      ],
    });

    expect(() => send(cut())).toThrowError(cutUnderRuleMessage);
  });
});

const gadgetGid = 11;
const gadgetColumnConfigs = {
  size: { columnId: "g:size", header: "Size" },
  colour: { columnId: "g:colour", header: "Colour" },
  unused: { columnId: "g:unused", header: "Unused" },
};

function gadgetSheet(placement: FakeTablePlacement = {}): FakeSheetProperties {
  return fakeTableSheet.build({
    sheetId: gadgetGid,
    title: "Gadget",
    columnConfigs: gadgetColumnConfigs,
    columnNames: ["size", "colour"],
    bodyRows: [
      { size: 1, colour: "red" },
      { size: { value: 2, numberFormatType: "NUMBER" } },
    ],
    ...placement,
  });
}

describe("fakeTableSheet builds a one-Table sheet from the layout", () => {
  it("puts the column IDs and headers on their head rows and the body below", () => {
    const { grid } = stubSheetsService({ sheets: [gadgetSheet()] });

    const sheet = grid.sheet(gadgetGid);
    expect(sheet.title).toBe("Gadget");
    expect(sheet.values()[expectedSheetLayout.colIdRowIndex]).toEqual([
      "g:size",
      "g:colour",
    ]);
    expect(sheet.values()[expectedSheetLayout.tableHeaderRowIndex]).toEqual([
      "Size",
      "Colour",
    ]);
    expect(
      sheet.values({ startRowIndex: expectedSheetLayout.topDataRowIndex }),
    ).toEqual([
      [1, "red"],
      [2, null],
    ]);
    expect(sheet.tables[0]?.range).toMatchObject({
      startRowIndex: expectedSheetLayout.tableHeaderRowIndex,
      endRowIndex: expectedSheetLayout.topDataRowIndex + 2,
      startColumnIndex: 0,
      endColumnIndex: 2,
    });
  });

  it("moves the head rows and body with a Table placed lower and to the right", () => {
    const { grid } = stubSheetsService({
      sheets: [gadgetSheet({ startRowIndex: 6, startColumnIndex: 1 })],
    });

    const sheet = grid.sheet(gadgetGid);
    expect(sheet.values({ startRowIndex: 3 })).toEqual([
      [null, "g:size", "g:colour"],
      [null, null, null],
      [null, null, null],
      [null, "Size", "Colour"],
      [null, 1, "red"],
      [null, 2, null],
    ]);
    expect(sheet.tables[0]?.range).toMatchObject({
      startRowIndex: 6,
      endRowIndex: 9,
      startColumnIndex: 1,
      endColumnIndex: 3,
    });
  });

  it("carries a given Table ID, column facts and head rows onto the Table", () => {
    const { grid } = stubSheetsService({
      sheets: [
        fakeTableSheet.build({
          sheetId: gadgetGid,
          title: "Gadget",
          tableId: "gadget",
          columnConfigs: gadgetColumnConfigs,
          columnNames: ["size", "colour"],
          bodyRows: [{}],
          columnTypes: { 0: "DOUBLE" },
          columnValidationValues: { 1: ["=Paint[Colour]"] },
          columnValidationConditionTypes: { 1: "ONE_OF_RANGE" },
          headRows: { 1: [false, "Looks"] },
        }),
      ],
    });

    const sheet = grid.sheet(gadgetGid);
    expect(sheet.tables[0]?.tableId).toBe("gadget");
    expect(sheet.tables[0]?.columnProperties).toEqual([
      { columnName: "Size", columnType: "DOUBLE" },
      {
        columnIndex: 1,
        columnName: "Colour",
        dataValidationRule: {
          condition: {
            type: "ONE_OF_RANGE",
            values: [{ userEnteredValue: "=Paint[Colour]" }],
          },
        },
      },
    ]);
    expect(sheet.values()[expectedSheetLayout.tableHeaderRowIndex - 1]).toEqual(
      [false, "Looks"],
    );
  });
});

describe("FakeSheetView reads a Table's body", () => {
  it("counts rows from the first body row and columns from the Table's first column", () => {
    const { grid } = stubSheetsService({
      sheets: [gadgetSheet({ startRowIndex: 6, startColumnIndex: 1 })],
    });

    const sheet = grid.sheet(gadgetGid);
    expect(sheet.bodyValues()).toEqual([
      [1, "red"],
      [2, null],
    ]);
    expect(sheet.bodyRows({ startRowIndex: 1, endColumnIndex: 1 })).toEqual([
      [{ value: 2, numberFormatType: "NUMBER" }],
    ]);
  });

  it("keeps a blank body row at the Table's end", () => {
    const { grid } = stubSheetsService({
      sheets: [
        fakeTableSheet.build({
          sheetId: gadgetGid,
          title: "Gadget",
          columnConfigs: gadgetColumnConfigs,
          columnNames: ["size"],
          bodyRows: [{ size: 1 }, {}],
        }),
      ],
    });

    expect(grid.sheet(gadgetGid).bodyValues()).toEqual([[1], [null]]);
  });

  it("reaches past the Table to a cell a write left below or beside it", () => {
    const fixture = gadgetSheet();
    const rows = (fixture.rows ?? []).map((row) => [...row]);
    rows.push([], [null, null, "beside"]);
    const { grid } = stubSheetsService({ sheets: [{ ...fixture, rows }] });

    expect(grid.sheet(gadgetGid).bodyValues()).toEqual([
      [1, "red", null],
      [2, null, null],
      [null, null, null],
      [null, null, "beside"],
    ]);
  });

  it("throws on a sheet that doesn't hold exactly one Table", () => {
    const { grid } = stubSheetsService({
      sheets: [bandSheet({ lower: lowerTable(1) })],
    });

    expect(() => grid.sheet(bandGid).bodyRows()).toThrowError(
      /Band holds 2 Tables; a body read needs exactly one/,
    );
  });
});

describe("stubSheetsService as a whole", () => {
  it("throws naming a request kind it does not replay", () => {
    stubSheetsService({ sheets: [widgetSheet()] });

    expect(() =>
      send({
        mergeCells: { range: { sheetId: widgetGid }, mergeType: "MERGE_ALL" },
      }),
    ).toThrowError("The fake Sheets service does not replay mergeCells.");
  });

  it("throws deleting a protected range no sheet holds, as live", () => {
    stubSheetsService({ sheets: [widgetSheet()] });

    expect(() =>
      send({ deleteProtectedRange: { protectedRangeId: 404 } }),
    ).toThrowError(/no protected range with id 404/);
  });

  it("applies a batch all or nothing", () => {
    const { grid } = stubSheetsService({ sheets: [widgetSheet()] });

    expect(() =>
      send(
        {
          updateSheetProperties: {
            properties: { sheetId: widgetGid, title: "Lost" },
            fields: "title",
          },
        },
        { deleteSheet: { sheetId: 404 } },
      ),
    ).toThrowError(/no sheet with gid 404/);
    expect(grid.sheetTitles()).toEqual(["Widget"]);
  });

  it("counts batch updates, and in a dry run applies none", () => {
    const { grid, batchUpdateCount } = stubSheetsService({
      sheets: [widgetSheet()],
      isDryRun: true,
    });

    send({ deleteSheet: { sheetId: widgetGid } });
    send({ deleteSheet: { sheetId: widgetGid } });

    expect(batchUpdateCount()).toBe(2);
    expect(grid.sheetTitles()).toEqual(["Widget"]);
  });

  it.each([
    {
      label:
        "gives a non-formula value an empty userEnteredValue under a formula-only mask",
      cellFields: "effectiveValue,userEnteredValue(formulaValue)",
      userEnteredValues: [{}, { formulaValue: "=1+1" }, undefined],
    },
    {
      label:
        "leaves a non-formula value's userEnteredValue out when the mask names no formula",
      cellFields: "effectiveValue",
      userEnteredValues: [undefined, { formulaValue: "=1+1" }, undefined],
    },
  ])("$label", ({ cellFields, userEnteredValues }) => {
    const { sheets } = stubSheetsService({
      sheets: [
        {
          sheetId: widgetGid,
          title: "Widget",
          rows: [["x", { value: 2, isFormula: true, formula: "=1+1" }, null]],
        },
      ],
    });

    const read = sheets.Spreadsheets.getByDataFilter(
      { dataFilters: [{ gridRange: { sheetId: widgetGid } }] },
      "any",
      { fields: `sheets(data(rowData(values(${cellFields}))))` },
    );

    const cells = read.sheets?.[0]?.data?.[0]?.rowData?.[0]?.values ?? [];
    expect(cells.map((cell) => cell.userEnteredValue)).toEqual(
      userEnteredValues,
    );
  });

  it("leaves a fixture untouched, so a test's writes never reach the next test that shares it", () => {
    const shared = widgetSheet();
    stubSheetsService({ sheets: [shared] });

    send({ deleteDimension: { range: rowRange(1, 2) } });

    expect(shared).toEqual(widgetSheet());
  });
});
