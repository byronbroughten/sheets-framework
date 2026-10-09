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

  it("leaves a fixture untouched, so a test's writes never reach the next test that shares it", () => {
    const shared = widgetSheet();
    stubSheetsService({ sheets: [shared] });

    send({ deleteDimension: { range: rowRange(1, 2) } });

    expect(shared).toEqual(widgetSheet());
  });
});
