import { describe, expect, it } from "vitest";

import { fakeSheetsHttpTransport } from "./fakeSheetsHttpTransport.ts";
import { stubSheetsService } from "./fakeSheetsService.ts";

type Request = GoogleAppsScript.Sheets.Schema.Request;

const base =
  "https://sheets.googleapis.com/v4/spreadsheets/spreadsheet-under-test";
const widgetGid = 7;

function seed() {
  const fake = stubSheetsService({
    sheets: [
      {
        sheetId: widgetGid,
        title: "Widget",
        rows: [
          ["a", "b"],
          [1, 2],
        ],
      },
    ],
  });
  const transport = fakeSheetsHttpTransport(fake);
  function post(body: object) {
    return transport({
      method: "POST",
      url: `${base}:batchUpdate`,
      body: JSON.stringify(body),
    });
  }
  function readSheet() {
    return transport({
      method: "GET",
      url: `${base}/values:batchGet?ranges=Widget!A:${lastColumn()}&valueRenderOption=UNFORMATTED_VALUE`,
      body: null,
    });
  }
  function lastColumn() {
    return "ABCDEFGH"[fake.grid.sheet(widgetGid).columnCount - 1];
  }
  return { fake, transport, post, readSheet };
}

function columns(startIndex: number, endIndex: number) {
  return { sheetId: widgetGid, dimension: "COLUMNS", startIndex, endIndex };
}

function columnBand(startColumnIndex: number, endColumnIndex: number) {
  return {
    sheetId: widgetGid,
    startRowIndex: 0,
    endRowIndex: 2,
    startColumnIndex,
    endColumnIndex,
  };
}

// Writes both rows of one column.
function write(columnIndex: number, top: string, bottom = top): Request {
  return {
    updateCells: {
      range: columnBand(columnIndex, columnIndex + 1),
      rows: [top, bottom].map((stringValue) => ({
        values: [{ userEnteredValue: { stringValue } }],
      })),
      fields: "userEnteredValue",
    },
  };
}

describe("fakeSheetsHttpTransport batchUpdate", () => {
  const cases: [string, Request[], (string | number)[][]][] = [
    [
      "an appended column",
      [
        {
          appendDimension: {
            sheetId: widgetGid,
            dimension: "COLUMNS",
            length: 1,
          },
        },
        write(2, "c", "3"),
      ],
      [
        ["a", "b", "c"],
        [1, 2, "3"],
      ],
    ],
    [
      "an inserted dimension",
      [{ insertDimension: { range: columns(1, 2) } }],
      [
        ["a", "", "b"],
        [1, "", 2],
      ],
    ],
    [
      "a deleted dimension",
      [{ deleteDimension: { range: columns(0, 1) } }],
      [["b"], [2]],
    ],
    [
      "a column inserted as a range",
      [
        {
          insertRange: { range: columnBand(0, 1), shiftDimension: "COLUMNS" },
        },
        write(0, "new", "0"),
      ],
      [
        ["new", "a", "b"],
        ["0", 1, 2],
      ],
    ],
    [
      "a column deleted as a range",
      [
        {
          deleteRange: { range: columnBand(0, 1), shiftDimension: "COLUMNS" },
        },
      ],
      [["b"], [2]],
    ],
  ];

  it.each(cases)(
    "shows %s in the grid and in a later batchGet",
    (_name, requests, expected) => {
      const { fake, post, readSheet } = seed();

      post({ requests });

      expect(
        fake.grid
          .sheet(widgetGid)
          .values()
          .map((row) => row.map((value) => value ?? "")),
      ).toMatchObject(expected);
      expect(readSheet()).toMatchObject({
        valueRanges: [{ values: expected }],
      });
    },
  );

  it("answers one reply for each request", () => {
    const { post } = seed();

    expect(post({ requests: [write(0, "x"), write(1, "y")] })).toEqual({
      replies: [{}, {}],
    });
  });
});

describe("fakeSheetsHttpTransport writes it doesn't model", () => {
  const url = `${base}:batchUpdate`;

  it("throws on a request kind with no replay, leaving the whole batch unapplied", () => {
    const { fake, post } = seed();
    const requests = [write(0, "x"), { autoResizeDimensions: {} }];

    expect(() => post({ requests })).toThrow(`POST ${url}`);
    expect(() => post({ requests })).toThrow("autoResizeDimensions");
    expect(fake.grid.sheet(widgetGid).cell(0, 0)).toBe("a");
  });

  it("keeps Google's refusal text in the error", () => {
    const { post } = seed();
    const pastGrid = {
      insertRange: { range: columnBand(5, 6), shiftDimension: "COLUMNS" },
    };

    expect(() => post({ requests: [pastGrid] })).toThrow(
      "Invalid requests[0].insertRange: Range",
    );
  });

  it.each([
    [
      "a body key beyond requests",
      url,
      '{"includeSpreadsheetInResponse":true}',
    ],
    ["a query parameter", `${url}?fields=replies`, '{"requests":[]}'],
    ["no body", url, null],
  ])("throws on %s, naming the method and URL", (_name, requestUrl, body) => {
    const { transport } = seed();

    expect(() => transport({ method: "POST", url: requestUrl, body })).toThrow(
      `POST ${requestUrl}`,
    );
  });
});
