import { describe, expect, it } from "vitest";

import { googleRawRequest } from "../00_Source/GoogleSheets/GoogleSheetsAPI";
import { installedRawSource } from "../00_Source/RawSource/RawSource";
import {
  buildGridRows,
  type FakeSheetProperties,
  stubSheetsService,
} from "./fakeSheetsService";

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
