import { describe, expect, it } from "vitest";

import { fakeSheetsHttpTransport } from "./fakeSheetsHttpTransport.ts";
import { buildGridRows, stubSheetsService } from "./fakeSheetsService.ts";

const base =
  "https://sheets.googleapis.com/v4/spreadsheets/spreadsheet-under-test";
const widgetGid = 7;

function seed() {
  const fake = stubSheetsService({
    sheets: [
      {
        sheetId: widgetGid,
        title: "Widget",
        rows: buildGridRows({
          0: ["Name", "Qty", "Total"],
          1: ["Bolt", 4, { value: 8, isFormula: true, formula: "=B2*2" }],
          3: [null, 0, ""],
          4: [],
        }),
        table: { startRowIndex: 0, startColumnIndex: 0, endRowIndex: 4 },
        conditionalFormats: [{ ranges: [{ sheetId: widgetGid }] }],
      },
      { sheetId: 8, title: "Bo's list", rows: [["only"]] },
    ],
  });
  return { fake, transport: fakeSheetsHttpTransport(fake) };
}

function get(query: string): unknown {
  return seed().transport({ method: "GET", url: base + query, body: null });
}

function batchGet(...params: string[]): unknown {
  return get(`/values:batchGet?${params.join("&")}`);
}

function ranges(...a1Ranges: string[]): string[] {
  return a1Ranges.map((range) => `ranges=${encodeURIComponent(range)}`);
}

function getByDataFilter(startRowIndex: number): unknown {
  return seed().transport({
    method: "POST",
    url: `${base}:getByDataFilter`,
    body: JSON.stringify({
      dataFilters: [{ gridRange: { sheetId: widgetGid, startRowIndex } }],
      includeGridData: true,
    }),
  });
}

const widgetTable = {
  range: { startRowIndex: 0, endRowIndex: 4 },
};

describe("fakeSheetsHttpTransport spreadsheet reads", () => {
  it("answers a get with the seeded grid, its Tables and its rules", () => {
    expect(get("")).toMatchObject({
      properties: { timeZone: "America/Chicago" },
      sheets: [
        {
          properties: {
            sheetId: widgetGid,
            title: "Widget",
            gridProperties: { rowCount: 5, columnCount: 3 },
          },
          tables: [widgetTable],
          conditionalFormats: [{ ranges: [{ sheetId: widgetGid }] }],
          data: [
            {
              rowData: [
                {
                  values: [{ effectiveValue: { stringValue: "Name" } }, {}, {}],
                },
                {
                  values: [
                    {},
                    { effectiveValue: { numberValue: 4 } },
                    { userEnteredValue: { formulaValue: "=B2*2" } },
                  ],
                },
                {},
                {},
                {},
              ],
            },
          ],
        },
        { properties: { sheetId: 8, title: "Bo's list" } },
      ],
    });
  });

  it("applies a get's encoded field mask", () => {
    const response = get(`?fields=${encodeURIComponent("sheets(properties)")}`);

    expect(response).not.toHaveProperty("properties");
  });

  it("answers getByDataFilter with the Tables a filter overlaps and no rules", () => {
    expect(getByDataFilter(0)).toMatchObject({
      sheets: [{ tables: [widgetTable] }, {}],
    });
    expect(getByDataFilter(0)).not.toHaveProperty(
      "sheets.0.conditionalFormats",
    );
    expect(getByDataFilter(4)).not.toHaveProperty("sheets.0.tables.0");
  });
});

describe("fakeSheetsHttpTransport values:batchGet", () => {
  it("answers each range in order, trimming trailing empty cells and rows", () => {
    expect(
      batchGet(
        ...ranges("Widget!A1:C5", "Widget!C:C", "Widget!A5:A", "Widget!B2:B2"),
        "valueRenderOption=UNFORMATTED_VALUE",
        "dateTimeRenderOption=SERIAL_NUMBER",
      ),
    ).toEqual({
      spreadsheetId: "spreadsheet-under-test",
      valueRanges: [
        {
          range: "Widget!A1:C5",
          majorDimension: "ROWS",
          values: [["Name", "Qty", "Total"], ["Bolt", 4, 8], [], ["", 0]],
        },
        {
          range: "Widget!C1:C5",
          majorDimension: "ROWS",
          values: [["Total"], [8]],
        },
        { range: "Widget!A5", majorDimension: "ROWS" },
        { range: "Widget!B2", majorDimension: "ROWS", values: [[4]] },
      ],
    });
  });

  it("answers by column, with a formula's text under FORMULA", () => {
    expect(
      batchGet(
        ...ranges("Widget!A:C"),
        "majorDimension=COLUMNS",
        "valueRenderOption=FORMULA",
      ),
    ).toMatchObject({
      valueRanges: [
        {
          range: "Widget!A1:C5",
          majorDimension: "COLUMNS",
          values: [
            ["Name", "Bolt"],
            ["Qty", 4, "", 0],
            ["Total", "=B2*2"],
          ],
        },
      ],
    });
  });

  it("reads a quoted title, a plus standing for a space", () => {
    expect(
      batchGet("ranges='Bo''s+list'!A1", "valueRenderOption=UNFORMATTED_VALUE"),
    ).toMatchObject({
      valueRanges: [{ range: "'Bo''s list'!A1", values: [["only"]] }],
    });
  });
});

describe("fakeSheetsHttpTransport reads it doesn't model", () => {
  const rendered = "valueRenderOption=UNFORMATTED_VALUE";

  it.each([
    ["no range", [rendered], "at least one range"],
    ["the default rendering", ranges("Widget!A1"), "FORMATTED_VALUE"],
    [
      "formatted dates",
      [
        ...ranges("Widget!A1"),
        rendered,
        "dateTimeRenderOption=FORMATTED_STRING",
      ],
      "FORMATTED_STRING",
    ],
    [
      "another major dimension",
      [...ranges("Widget!A1"), rendered, "majorDimension=SHEETS"],
      "SHEETS",
    ],
    [
      "an unknown parameter",
      [...ranges("Widget!A1"), rendered, "fields=range"],
      '"fields"',
    ],
    ["whole rows", [...ranges("Widget!1:2"), rendered], "Widget!1:2"],
    ["a bare sheet name", [...ranges("Widget"), rendered], '"Widget"'],
    ["an unknown title", [...ranges("Gadget!A1"), rendered], "Gadget"],
    ["a row past the grid", [...ranges("Widget!A1:A6"), rendered], "grid"],
    ["a column past the grid", [...ranges("Widget!D:D"), rendered], "grid"],
  ])("throws on %s, naming the method and URL", (_name, params, detail) => {
    const url = `${base}/values:batchGet?${params.join("&")}`;

    expect(() => batchGet(...params)).toThrow(`GET ${url}`);
    expect(() => batchGet(...params)).toThrow(detail);
  });

  it.each([
    ["GET", ":export", null],
    ["GET", ":batchUpdate", null],
    ["POST", "", "{}"],
    ["GET", "?includeGridData=true", null],
    ["POST", ":getByDataFilter", null],
    ["POST", ":getByDataFilter", '{"excludeTablesInBandedRanges":true}'],
    ["POST", ":getByDataFilter", '{"dataFilters":[{"a1Range":"Widget!A1"}]}'],
  ] as const)(
    "throws on %s %s with body %s, naming the method and URL",
    (method, suffix, body) => {
      const url = base + suffix;

      expect(() => seed().transport({ method, url, body })).toThrow(
        `${method} ${url}`,
      );
    },
  );

  it("throws on a URL outside the Sheets API", () => {
    const url = "https://example.invalid/v4/spreadsheets/x";

    expect(() => seed().transport({ method: "GET", url, body: null })).toThrow(
      `GET ${url}`,
    );
  });
});
