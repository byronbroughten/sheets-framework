import { describe, expect, it, vi } from "vitest";

import { stubScriptAndSpreadsheetApp } from "../../testSupport/fakeAppsScriptGlobals";
import type {
  AddCheckboxValidationOperation,
  AddSheetOperation,
  AddTableOperation,
  LocalWriteOperation,
} from "../RawSource/RawSource";
import type { RgbColor } from "../RawSource/RgbColor";
import {
  googleRawRequest,
  type GoogleRequest,
  GoogleSheetsAPI,
  type SheetsHttpRequest,
} from "./GoogleSheetsAPI";

const spreadsheetId = "spreadsheet-under-test";
const lightGreen: RgbColor = { red: 0.851, green: 0.918, blue: 0.827 };

const addSheetOperation: AddSheetOperation = {
  kind: "addSheet",
  sheetId: 555,
  title: "Spreadsheet Config",
  rowCount: 20,
  columnCount: 6,
};
const addTableOperation: AddTableOperation = {
  kind: "addTable",
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
    { columnIndex: 2, columnName: "Amount", columnType: "CURRENCY" },
  ],
};
const checkboxValidationOperation: AddCheckboxValidationOperation = {
  kind: "addCheckboxValidation",
  range: {
    sheetId: 555,
    startRowIndex: 1,
    endRowIndex: 2,
    startColumnIndex: 3,
    endColumnIndex: 4,
  },
};

type BatchUpdateRequest =
  GoogleAppsScript.Sheets.Schema.BatchUpdateSpreadsheetRequest;

function recordingSheets(
  payload: GoogleAppsScript.Sheets.Schema.Spreadsheet = {
    sheets: [],
  },
  batchUpdateResponse: (
    resource: BatchUpdateRequest,
  ) => GoogleAppsScript.Sheets.Schema.BatchUpdateSpreadsheetResponse = () => ({}),
) {
  const batchUpdateCalls: BatchUpdateRequest[] = [];
  const getByDataFilterCalls: object[] = [];
  const getByDataFilterFields: (string | undefined)[] = [];
  const getCalls: { spreadsheetId: string; fields?: string }[] = [];
  const requestedIds: string[] = [];
  const sheets = {
    Spreadsheets: {
      get: (spreadsheetId: string, optionalArgs?: { fields?: string }) => {
        requestedIds.push(spreadsheetId);
        getCalls.push({ spreadsheetId, fields: optionalArgs?.fields });
        return payload;
      },
      getByDataFilter: (
        resource: object,
        spreadsheetId: string,
        optionalArgs?: { fields?: string },
      ) => {
        requestedIds.push(spreadsheetId);
        getByDataFilterCalls.push(resource);
        getByDataFilterFields.push(optionalArgs?.fields);
        return payload;
      },
      batchUpdate: (resource: BatchUpdateRequest, spreadsheetId: string) => {
        requestedIds.push(spreadsheetId);
        batchUpdateCalls.push(resource);
        return batchUpdateResponse(resource);
      },
    },
  };
  return {
    api: GoogleSheetsAPI.init(sheets, spreadsheetId),
    batchUpdateCalls,
    getByDataFilterCalls,
    getByDataFilterFields,
    getCalls,
    requestedIds,
  };
}

describe("GoogleSheetsAPI write mapping", () => {
  it("maps each local operation kind onto the Google request used today", () => {
    const { api, batchUpdateCalls } = recordingSheets();

    const operations: LocalWriteOperation[] = [
      {
        kind: "appendRows",
        sheetId: 111,
        tableId: "tbl",
        emptyRowCount: 2,
      },
      { kind: "insertColumn", sheetId: 111, startColumnIndex: 3 },
      {
        kind: "fill",
        sheetId: 111,
        colIndex: 2,
        startRowIndex: 4,
        endRowIndex: 6,
        value: "x",
        backgroundColor: lightGreen,
      },
      {
        kind: "fill",
        sheetId: 111,
        colIndex: 2,
        startRowIndex: 4,
        endRowIndex: 6,
        formula: "=A4",
      },
      {
        kind: "updateCell",
        sheetId: 111,
        rowIndex: 5,
        colIndex: 2,
        value: "y",
      },
      {
        kind: "updateCell",
        sheetId: 111,
        rowIndex: 5,
        colIndex: 2,
        formula: "=B5",
        backgroundColor: lightGreen,
      },
      {
        kind: "findReplace",
        terms: { find: "a", replacement: "b" },
        scope: { sheetId: 111 },
      },
      { kind: "deleteRows", sheetId: 111, startIndex: 8, endIndex: 9 },
      {
        kind: "sort",
        sheetId: 111,
        startRowIndex: 4,
        startColumnIndex: 0,
        colIdxToSortBy: 1,
        sortOrder: "ASCENDING",
      },
      {
        kind: "addProtectedRange",
        protection: {
          kind: "warning",
          range: {
            sheetId: 111,
            startRowIndex: 4,
            endRowIndex: 5,
            startColumnIndex: 0,
            endColumnIndex: 1,
          },
          description: "floor",
          users: [],
          groups: [],
          unprotectedRanges: [],
        },
      },
      { kind: "deleteProtectedRange", sheetId: 111, protectedRangeId: 7 },
      {
        kind: "updateSheetTitle",
        sheetId: 111,
        title: "Spreadsheet Config",
      },
      { kind: "updateTableName", tableId: "tbl", name: "spreadsheetConfig" },
      {
        kind: "updateTableColumnProperties",
        tableId: "tbl",
        columnProperties: [
          {
            columnIndex: 2,
            columnName: "Amount",
            columnType: "TEXT",
          },
        ],
      },
      {
        kind: "raw",
        request: googleRawRequest({ updateTable: { table: { tableId: "t" } } }),
      },
    ];

    api.flush(operations);

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        appendCells: {
          sheetId: 111,
          tableId: "tbl",
          rows: [{}, {}],
          fields: "userEnteredValue",
        },
      },
      {
        insertDimension: {
          range: {
            sheetId: 111,
            dimension: "COLUMNS",
            startIndex: 3,
            endIndex: 4,
          },
          inheritFromBefore: false,
        },
      },
      {
        repeatCell: {
          range: {
            sheetId: 111,
            startRowIndex: 4,
            endRowIndex: 6,
            startColumnIndex: 2,
            endColumnIndex: 3,
          },
          cell: {
            userEnteredValue: { stringValue: "x" },
            userEnteredFormat: { backgroundColor: lightGreen },
          },
          fields: "userEnteredValue,userEnteredFormat.backgroundColor",
        },
      },
      {
        pasteData: {
          coordinate: { sheetId: 111, rowIndex: 4, columnIndex: 2 },
          data: `"=A4"\n"=A4"`,
          delimiter: "\t",
          type: "PASTE_FORMULA",
        },
      },
      {
        updateCells: {
          range: {
            sheetId: 111,
            startRowIndex: 5,
            endRowIndex: 6,
            startColumnIndex: 2,
            endColumnIndex: 3,
          },
          rows: [{ values: [{ userEnteredValue: { stringValue: "y" } }] }],
          fields: "userEnteredValue",
        },
      },
      {
        pasteData: {
          coordinate: { sheetId: 111, rowIndex: 5, columnIndex: 2 },
          data: `"=B5"`,
          delimiter: "\t",
          type: "PASTE_FORMULA",
        },
      },
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
      { findReplace: { find: "a", replacement: "b", sheetId: 111 } },
      {
        deleteDimension: {
          range: {
            sheetId: 111,
            dimension: "ROWS",
            startIndex: 8,
            endIndex: 9,
          },
        },
      },
      {
        sortRange: {
          range: { sheetId: 111, startRowIndex: 4, startColumnIndex: 0 },
          sortSpecs: [{ dimensionIndex: 1, sortOrder: "ASCENDING" }],
        },
      },
      {
        addProtectedRange: {
          protectedRange: {
            range: {
              sheetId: 111,
              startRowIndex: 4,
              endRowIndex: 5,
              startColumnIndex: 0,
              endColumnIndex: 1,
            },
            description: "floor",
            warningOnly: true,
          },
        },
      },
      { deleteProtectedRange: { protectedRangeId: 7 } },
      {
        updateSheetProperties: {
          properties: { sheetId: 111, title: "Spreadsheet Config" },
          fields: "title",
        },
      },
      {
        updateTable: {
          table: { tableId: "tbl", name: "spreadsheetConfig" },
          fields: "name",
        },
      },
      {
        updateTable: {
          table: {
            tableId: "tbl",
            columnProperties: [
              { columnIndex: 2, columnName: "Amount", columnType: "TEXT" },
            ],
          },
          fields: "columnProperties",
        },
      },
      { updateTable: { table: { tableId: "t" } } },
    ]);
  });

  it("sends no batchUpdate when the write list is empty", () => {
    const { api, batchUpdateCalls } = recordingSheets();

    api.flush([]);

    expect(batchUpdateCalls).toEqual([]);
  });

  it("maps a boolean rule onto addConditionalFormatRule without alpha", () => {
    const { api, batchUpdateCalls } = recordingSheets();
    const pink = { red: 244 / 255, green: 204 / 255, blue: 204 / 255 };

    api.flush([
      {
        kind: "addConditionalFormatRule",
        index: 0,
        rule: {
          kind: "boolean",
          ranges: [
            {
              sheetId: 111,
              startRowIndex: 4,
              endRowIndex: 11,
              startColumnIndex: 2,
              endColumnIndex: 3,
            },
          ],
          condition: { type: "NUMBER_EQ", value: true },
          format: {
            backgroundColor: pink,
            foregroundColor: { red: 0.4, green: 0.4, blue: 0.4 },
          },
        },
      },
    ]);

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        addConditionalFormatRule: {
          index: 0,
          rule: {
            ranges: [
              {
                sheetId: 111,
                startRowIndex: 4,
                endRowIndex: 11,
                startColumnIndex: 2,
                endColumnIndex: 3,
              },
            ],
            booleanRule: {
              condition: {
                type: "NUMBER_EQ",
                values: [{ userEnteredValue: "TRUE" }],
              },
              format: {
                backgroundColor: pink,
                textFormat: {
                  foregroundColor: { red: 0.4, green: 0.4, blue: 0.4 },
                },
              },
            },
          },
        },
      },
    ]);
    expect(
      batchUpdateCalls[0]?.requests?.[0]?.addConditionalFormatRule?.rule
        ?.booleanRule?.format?.backgroundColor,
    ).not.toHaveProperty("alpha");
  });

  it("stringifies a false condition value as Sheets' FALSE literal", () => {
    const { api, batchUpdateCalls } = recordingSheets();

    api.flush([
      {
        kind: "addConditionalFormatRule",
        index: 0,
        rule: {
          kind: "boolean",
          ranges: [
            {
              sheetId: 111,
              startRowIndex: 4,
              endRowIndex: 5,
              startColumnIndex: 0,
              endColumnIndex: 1,
            },
          ],
          condition: { type: "NUMBER_EQ", value: false },
          format: { backgroundColor: { red: 1 } },
        },
      },
    ]);

    expect(
      batchUpdateCalls[0]?.requests?.[0]?.addConditionalFormatRule?.rule
        ?.booleanRule?.condition?.values,
    ).toEqual([{ userEnteredValue: "FALSE" }]);
  });

  it("passes a custom formula through unchanged and maps a delete by sheet and index", () => {
    const { api, batchUpdateCalls } = recordingSheets();

    api.flush([
      {
        kind: "deleteConditionalFormatRule",
        sheetId: 111,
        index: 3,
      },
      {
        kind: "addConditionalFormatRule",
        index: 0,
        rule: {
          kind: "boolean",
          ranges: [
            {
              sheetId: 111,
              startRowIndex: 4,
              endRowIndex: 5,
              startColumnIndex: 0,
              endColumnIndex: 1,
            },
          ],
          condition: { type: "CUSTOM_FORMULA", formula: "=$B5=FALSE" },
          format: { backgroundColor: { red: 1 } },
        },
      },
    ]);

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        deleteConditionalFormatRule: { sheetId: 111, index: 3 },
      },
      {
        addConditionalFormatRule: {
          index: 0,
          rule: {
            ranges: [
              {
                sheetId: 111,
                startRowIndex: 4,
                endRowIndex: 5,
                startColumnIndex: 0,
                endColumnIndex: 1,
              },
            ],
            booleanRule: {
              condition: {
                type: "CUSTOM_FORMULA",
                values: [{ userEnteredValue: "=$B5=FALSE" }],
              },
              format: { backgroundColor: { red: 1 } },
            },
          },
        },
      },
    ]);
  });

  it("passes an ordered raw request through last", () => {
    const { api, batchUpdateCalls } = recordingSheets();

    api.flush([
      {
        kind: "updateCell",
        sheetId: 111,
        rowIndex: 1,
        colIndex: 0,
        value: "a",
      },
      {
        kind: "raw",
        request: googleRawRequest({ updateTable: { table: { tableId: "t" } } }),
      },
    ]);

    expect(
      (batchUpdateCalls[0]?.requests ?? []).map(
        (request) => Object.keys(request)[0],
      ),
    ).toEqual(["updateCells", "updateTable"]);
  });

  it("maps an add-sheet operation onto one addSheet request carrying its sheetId, title and grid size", () => {
    const { api, batchUpdateCalls } = recordingSheets();

    api.flush([addSheetOperation]);

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
    ]);
  });

  it("maps an add-Table operation onto an addTable keyed by its name, then an updateTable carrying its columns", () => {
    const { api, batchUpdateCalls } = recordingSheets();

    api.flush([addTableOperation]);

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        addTable: {
          table: {
            tableId: "spreadsheetConfig",
            name: "spreadsheetConfig",
            range: addTableOperation.range,
          },
        },
      },
      {
        updateTable: {
          table: {
            tableId: "spreadsheetConfig",
            columnProperties: addTableOperation.columnProperties,
          },
          fields: "columnProperties",
        },
      },
    ]);
  });

  it("maps an add-sheet and an add-Table handed together onto three requests in that order, and nothing else", () => {
    const { api, batchUpdateCalls } = recordingSheets();

    api.flush([addSheetOperation, addTableOperation]);

    expect(batchUpdateCalls).toHaveLength(1);
    expect(
      (batchUpdateCalls[0]?.requests ?? []).map((request) =>
        Object.keys(request),
      ),
    ).toEqual([["addSheet"], ["addTable"], ["updateTable"]]);
  });

  it("maps a checkbox-validation operation onto one setDataValidation with a BOOLEAN condition over its range", () => {
    const { api, batchUpdateCalls } = recordingSheets();

    api.flush([checkboxValidationOperation]);

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        setDataValidation: {
          range: checkboxValidationOperation.range,
          rule: { condition: { type: "BOOLEAN" } },
        },
      },
    ]);
  });

  it("keeps a checkbox validation after the add-sheet it is handed behind", () => {
    const { api, batchUpdateCalls } = recordingSheets();

    api.flush([addSheetOperation, checkboxValidationOperation]);

    expect(
      (batchUpdateCalls[0]?.requests ?? []).map((request) =>
        Object.keys(request),
      ),
    ).toEqual([["addSheet"], ["setDataValidation"]]);
  });
});

describe("GoogleSheetsAPI payload mapping", () => {
  it("maps a Google spreadsheet payload onto the Raw-facing snapshot", () => {
    const { api } = recordingSheets({
      sheets: [
        {
          properties: { sheetId: 111, title: "Records" },
          tables: [
            {
              tableId: "tbl",
              range: {
                startRowIndex: 3,
                endRowIndex: 11,
                startColumnIndex: 0,
                endColumnIndex: 5,
              },
              columnProperties: [
                {
                  columnIndex: 1,
                  columnType: "CURRENCY",
                  dataValidationRule: {
                    condition: {
                      type: "ONE_OF_LIST",
                      values: [{ userEnteredValue: "Open" }],
                    },
                  },
                },
              ],
            },
          ],
          data: [
            {
              startColumn: 0,
              startRow: 4,
              columnMetadata: [{}, {}],
              rowData: [
                {
                  values: [
                    { effectiveValue: { stringValue: "id-1" } },
                    {
                      effectiveValue: { numberValue: 12 },
                      userEnteredValue: { formulaValue: "=A4" },
                      effectiveFormat: { numberFormat: { type: "CURRENCY" } },
                      dataValidation: { condition: { type: "NUMBER_GREATER" } },
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    });

    expect(api.fetchSheetProperties()).toEqual({
      timeZone: null,
      sheets: [
        {
          sheetGid: 111,
          title: "Records",
          tables: [
            {
              tableId: "tbl",
              name: "",
              startRowIndex: 3,
              endRowIndex: 11,
              startColumnIndex: 0,
              endColumnIndex: 5,
              columnProperties: [
                {
                  columnIndex: 1,
                  columnType: "CURRENCY",
                  dataValidationValues: ["Open"],
                  dataValidationConditionType: "ONE_OF_LIST",
                },
              ],
            },
          ],
          gridBlocks: [
            {
              startColumn: 0,
              startRow: 4,
              columnCount: 2,
              rows: [
                {
                  cells: [
                    {
                      value: "id-1",
                      isFormula: false,
                      numberFormatType: undefined,
                      dataValidationConditionType: undefined,
                    },
                    {
                      value: 12,
                      isFormula: true,
                      numberFormatType: "CURRENCY",
                      dataValidationConditionType: "NUMBER_GREATER",
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    });
  });

  it("reads a Table column with no columnIndex as index 0", () => {
    const { api } = recordingSheets({
      sheets: [
        {
          properties: { sheetId: 111, title: "Records" },
          tables: [
            {
              tableId: "tbl",
              range: {
                startRowIndex: 3,
                endRowIndex: 11,
                startColumnIndex: 2,
                endColumnIndex: 4,
              },
              columnProperties: [
                { columnName: "Name", columnType: "TEXT" },
                { columnIndex: 1, columnName: "Amount" },
              ],
            },
          ],
        },
      ],
    });

    const columns =
      api.fetchSheetProperties().sheets[0]?.tables?.[0]?.columnProperties;
    expect(columns?.map((column) => column.columnIndex)).toEqual([0, 1]);
  });

  it("maps a boolean rule, a multi-range rule and an unmodelable rule in list order", () => {
    const pink = { red: 0.95686275, green: 0.8, blue: 0.8 };
    const { api } = recordingSheets({
      sheets: [
        {
          properties: { sheetId: 111, title: "Records" },
          conditionalFormats: [
            {
              ranges: [
                {
                  sheetId: 111,
                  startRowIndex: 4,
                  endRowIndex: 11,
                  startColumnIndex: 0,
                  endColumnIndex: 5,
                },
              ],
              booleanRule: {
                condition: {
                  type: "CUSTOM_FORMULA",
                  values: [{ userEnteredValue: "=$A5" }],
                },
                format: {
                  backgroundColor: pink,
                  backgroundColorStyle: { rgbColor: pink },
                  textFormat: {
                    foregroundColor: { red: 0.4 },
                    foregroundColorStyle: { rgbColor: { red: 0.4 } },
                  },
                },
              },
            },
            {
              ranges: [
                {
                  sheetId: 111,
                  startRowIndex: 4,
                  endRowIndex: 6,
                  startColumnIndex: 0,
                  endColumnIndex: 1,
                },
                {
                  sheetId: 111,
                  startRowIndex: 4,
                  endRowIndex: 6,
                  startColumnIndex: 3,
                  endColumnIndex: 4,
                },
              ],
              booleanRule: {
                condition: {
                  type: "NUMBER_EQ",
                  values: [{ userEnteredValue: "TRUE" }],
                },
                format: { backgroundColor: { red: 0, green: 1, blue: 0 } },
              },
            },
            {
              ranges: [
                {
                  sheetId: 111,
                  startRowIndex: 4,
                  endRowIndex: 11,
                  startColumnIndex: 2,
                  endColumnIndex: 3,
                },
              ],
              gradientRule: {
                minpoint: { color: { red: 1 }, type: "MIN" },
                maxpoint: { color: { red: 0 }, type: "MAX" },
              },
            },
          ],
        },
      ],
    });

    const rules = api.fetchConditionalFormatRules()[0]?.rules;

    expect(rules).toHaveLength(3);
    expect(rules?.[0]).toEqual({
      kind: "boolean",
      ranges: [
        {
          sheetId: 111,
          startRowIndex: 4,
          endRowIndex: 11,
          startColumnIndex: 0,
          endColumnIndex: 5,
        },
      ],
      condition: { type: "CUSTOM_FORMULA", formula: "=$A5" },
      format: {
        backgroundColor: {
          red: 244 / 255,
          green: 204 / 255,
          blue: 204 / 255,
        },
        foregroundColor: { red: 102 / 255 },
      },
    });
    expect(rules?.[1]).toEqual({
      kind: "boolean",
      ranges: [
        {
          sheetId: 111,
          startRowIndex: 4,
          endRowIndex: 6,
          startColumnIndex: 0,
          endColumnIndex: 1,
        },
        {
          sheetId: 111,
          startRowIndex: 4,
          endRowIndex: 6,
          startColumnIndex: 3,
          endColumnIndex: 4,
        },
      ],
      condition: { type: "NUMBER_EQ", value: true },
      format: { backgroundColor: { red: 0, green: 1, blue: 0 } },
    });
    expect(rules?.[2]).toEqual({
      kind: "unmodelable",
      index: 2,
      ranges: [
        {
          sheetId: 111,
          startRowIndex: 4,
          endRowIndex: 11,
          startColumnIndex: 2,
          endColumnIndex: 3,
        },
      ],
    });
  });
});

describe("GoogleSheetsAPI time zone read", () => {
  const timeZoneMask = "properties(timeZone)";

  it("carries the spreadsheet's time zone in all three standing field masks", () => {
    const { api, getCalls, getByDataFilterFields } = recordingSheets();

    api.fetchSheetProperties();
    api.fetchGrid([{ sheetId: 1 }], { includeProgrammaticFacts: false });
    api.fetchGrid([{ sheetId: 1 }], { includeProgrammaticFacts: true });

    const masks = [getCalls[0]?.fields, ...getByDataFilterFields];
    expect(masks).toHaveLength(3);
    masks.forEach((mask) => {
      expect(mask?.startsWith(`${timeZoneMask},sheets(`)).toBe(true);
    });
  });

  it("maps the payload's time zone onto the snapshot", () => {
    const { api } = recordingSheets({
      properties: { timeZone: "Europe/Paris" },
      sheets: [],
    });

    expect(api.fetchSheetProperties().timeZone).toBe("Europe/Paris");
  });

  it("maps a payload with no time zone to null", () => {
    const { api } = recordingSheets({ sheets: [] });

    expect(api.fetchSheetProperties().timeZone).toBeNull();
  });

  it("fetches the time zone alone with a plain get masking only that field", () => {
    const { api, getCalls } = recordingSheets({
      properties: { timeZone: "Asia/Tokyo" },
    });

    expect(api.fetchTimeZone()).toBe("Asia/Tokyo");
    expect(getCalls).toEqual([{ spreadsheetId, fields: timeZoneMask }]);
  });

  it("answers null when the lone time zone fetch comes back without it", () => {
    const { api } = recordingSheets({});

    expect(api.fetchTimeZone()).toBeNull();
  });
});

describe("GoogleSheetsAPI conditional format read", () => {
  it("reads rules with a plain get, never getByDataFilter", () => {
    const { api, getCalls, getByDataFilterCalls } = recordingSheets({
      sheets: [{ properties: { sheetId: 111 } }],
    });

    api.fetchConditionalFormatRules();

    expect(getCalls).toEqual([
      {
        spreadsheetId,
        fields: "sheets(properties(sheetId),conditionalFormats)",
      },
    ]);
    expect(getByDataFilterCalls).toHaveLength(0);
  });

  it("reads omitted zero fields in a rule range as zero", () => {
    const { api } = recordingSheets({
      sheets: [
        {
          properties: { sheetId: 0 },
          conditionalFormats: [
            {
              ranges: [{ endRowIndex: 11, endColumnIndex: 1 }],
              gradientRule: {},
            },
          ],
        },
      ],
    });

    expect(api.fetchConditionalFormatRules()[0]?.rules[0]?.ranges).toEqual([
      {
        sheetId: 0,
        startRowIndex: 0,
        endRowIndex: 11,
        startColumnIndex: 0,
        endColumnIndex: 1,
      },
    ]);
  });

  it("reads a theme colour style, or one that disagrees with its colour, as unmodelable", () => {
    const green = { red: 0, green: 1, blue: 0 };
    const rule = (format: GoogleAppsScript.Sheets.Schema.CellFormat) => ({
      ranges: [{ sheetId: 111, startRowIndex: 4 }],
      booleanRule: {
        condition: {
          type: "NUMBER_EQ",
          values: [{ userEnteredValue: "TRUE" }],
        },
        format,
      },
    });
    const { api } = recordingSheets({
      sheets: [
        {
          properties: { sheetId: 111 },
          conditionalFormats: [
            rule({
              backgroundColor: green,
              backgroundColorStyle: { themeColor: "ACCENT1" },
            }),
            rule({
              backgroundColor: green,
              backgroundColorStyle: { rgbColor: { red: 1 } },
            }),
          ],
        },
      ],
    });

    expect(
      api.fetchConditionalFormatRules()[0]?.rules.map((read) => read.kind),
    ).toEqual(["unmodelable", "unmodelable"]);
  });

  it("reads a sheet whose rule list Google omitted as having no rules", () => {
    const { api } = recordingSheets({
      sheets: [{ properties: { sheetId: 111 } }],
    });

    expect(api.fetchConditionalFormatRules()).toEqual([
      { sheetGid: 111, rules: [] },
    ]);
  });
});

describe("GoogleSheetsAPI protected range read", () => {
  it("reads protections with a plain get, never getByDataFilter", () => {
    const { api, getCalls, getByDataFilterCalls } = recordingSheets({
      sheets: [{ properties: { sheetId: 111 } }],
    });

    api.fetchEditProtections();

    expect(getCalls).toEqual([
      {
        spreadsheetId,
        fields: "sheets(properties(sheetId),protectedRanges)",
      },
    ]);
    expect(getByDataFilterCalls).toHaveLength(0);
  });

  it("reads a sheet whose protection list Google omitted as having none", () => {
    const { api } = recordingSheets({
      sheets: [{ properties: { sheetId: 111 } }],
    });

    expect(api.fetchEditProtections()).toEqual([
      { sheetGid: 111, protections: [] },
    ]);
  });

  it("reads omitted zero fields in a protection range as zero", () => {
    const { api } = recordingSheets({
      sheets: [
        {
          properties: { sheetId: 0 },
          protectedRanges: [
            {
              protectedRangeId: 3,
              range: { endRowIndex: 11, endColumnIndex: 1 },
              warningOnly: true,
            },
          ],
        },
      ],
    });

    expect(api.fetchEditProtections()[0]?.protections[0]).toEqual({
      kind: "warning",
      id: 3,
      range: {
        sheetId: 0,
        startRowIndex: 0,
        endRowIndex: 11,
        startColumnIndex: 0,
        endColumnIndex: 1,
      },
      description: "",
      users: [],
      groups: [],
      unprotectedRanges: [],
      requestingUserCanEdit: false,
    });
  });

  it("reads a warning's editors as none, since Google lists them but a warning ignores them", () => {
    const { api } = recordingSheets({
      sheets: [
        {
          properties: { sheetId: 111 },
          protectedRanges: [
            {
              protectedRangeId: 5,
              range: {
                sheetId: 111,
                startRowIndex: 4,
                endRowIndex: 16,
                startColumnIndex: 1,
                endColumnIndex: 2,
              },
              warningOnly: true,
              editors: {
                users: [
                  "service@example.iam.gserviceaccount.com",
                  "owner@example.com",
                ],
                groups: ["team@example.com"],
              },
            },
          ],
        },
      ],
    });

    const protection = api.fetchEditProtections()[0]?.protections[0];
    expect(protection).toMatchObject({
      kind: "warning",
      users: [],
      groups: [],
    });
  });

  it("reads a named-range-backed protection as unmodelable and still carries its id", () => {
    const { api } = recordingSheets({
      sheets: [
        {
          properties: { sheetId: 111 },
          protectedRanges: [
            {
              protectedRangeId: 9,
              namedRangeId: "named-1",
              range: { sheetId: 111, startRowIndex: 4, endRowIndex: 11 },
              warningOnly: true,
            },
          ],
        },
      ],
    });

    expect(api.fetchEditProtections()[0]?.protections[0]).toEqual({
      kind: "unmodelable",
      id: 9,
    });
  });

  it("reads a protection with column bounds and no row bounds as a whole-column range, not a whole sheet", () => {
    const { api } = recordingSheets({
      sheets: [
        {
          properties: { sheetId: 111 },
          protectedRanges: [
            {
              protectedRangeId: 8,
              range: {
                sheetId: 111,
                startColumnIndex: 4,
                endColumnIndex: 5,
              },
              warningOnly: true,
            },
          ],
        },
      ],
    });

    expect(api.fetchEditProtections()[0]?.protections[0]).toEqual({
      kind: "warning",
      id: 8,
      range: {
        sheetId: 111,
        startRowIndex: 0,
        startColumnIndex: 4,
        endColumnIndex: 5,
      },
      description: "",
      users: [],
      groups: [],
      unprotectedRanges: [],
      requestingUserCanEdit: false,
    });
  });
});

describe("GoogleSheetsAPI protected range write", () => {
  it("maps a lock with named editors and a whole-sheet warning with unprotected ranges", () => {
    const { api, batchUpdateCalls } = recordingSheets();

    api.flush([
      {
        kind: "addProtectedRange",
        protection: {
          kind: "lock",
          range: {
            sheetId: 111,
            startRowIndex: 4,
            endRowIndex: 5,
            startColumnIndex: 0,
            endColumnIndex: 1,
          },
          description: "cell lock",
          users: ["editor@example.com"],
          groups: ["group@example.com"],
          unprotectedRanges: [],
        },
      },
      {
        kind: "addProtectedRange",
        protection: {
          kind: "warning",
          range: { sheetId: 111 },
          description: "sheet warning",
          users: [],
          groups: [],
          unprotectedRanges: [
            {
              sheetId: 111,
              startRowIndex: 4,
              endRowIndex: 5,
              startColumnIndex: 1,
              endColumnIndex: 2,
            },
          ],
        },
      },
    ]);

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        addProtectedRange: {
          protectedRange: {
            range: {
              sheetId: 111,
              startRowIndex: 4,
              endRowIndex: 5,
              startColumnIndex: 0,
              endColumnIndex: 1,
            },
            description: "cell lock",
            editors: {
              users: ["editor@example.com"],
              groups: ["group@example.com"],
            },
          },
        },
      },
      {
        addProtectedRange: {
          protectedRange: {
            range: { sheetId: 111 },
            description: "sheet warning",
            warningOnly: true,
            unprotectedRanges: [
              {
                sheetId: 111,
                startRowIndex: 4,
                endRowIndex: 5,
                startColumnIndex: 1,
                endColumnIndex: 2,
              },
            ],
          },
        },
      },
    ]);
  });

  it("writes a whole-column range as column bounds with no end row", () => {
    const { api, batchUpdateCalls } = recordingSheets();

    api.flush([
      {
        kind: "addProtectedRange",
        protection: {
          kind: "warning",
          range: {
            sheetId: 111,
            startRowIndex: 0,
            startColumnIndex: 4,
            endColumnIndex: 5,
          },
          description: "column warning",
          users: [],
          groups: [],
          unprotectedRanges: [],
        },
      },
    ]);

    expect(batchUpdateCalls[0]?.requests).toEqual([
      {
        addProtectedRange: {
          protectedRange: {
            range: {
              sheetId: 111,
              startRowIndex: 0,
              startColumnIndex: 4,
              endColumnIndex: 5,
            },
            description: "column warning",
            warningOnly: true,
          },
        },
      },
    ]);
  });

  it("reads the new protected range id from the add reply", () => {
    const { api } = recordingSheets({ sheets: [] }, () => ({
      replies: [
        {
          addProtectedRange: {
            protectedRange: { protectedRangeId: 42 },
          },
        },
      ],
    }));

    expect(() =>
      api.flush([
        {
          kind: "addProtectedRange",
          protection: {
            kind: "warning",
            range: { sheetId: 111 },
            description: "",
            users: [],
            groups: [],
            unprotectedRanges: [],
          },
        },
      ]),
    ).not.toThrow();
  });

  it("throws when an add reply omits the id", () => {
    const { api } = recordingSheets({ sheets: [] }, () => ({
      replies: [{ addProtectedRange: { protectedRange: {} } }],
    }));

    expect(() =>
      api.flush([
        {
          kind: "addProtectedRange",
          protection: {
            kind: "warning",
            range: { sheetId: 111 },
            description: "",
            users: [],
            groups: [],
            unprotectedRanges: [],
          },
        },
      ]),
    ).toThrowError(/protectedRangeId/);
  });
});

describe("GoogleSheetsAPI colour mapping", () => {
  it("maps the RGB record onto Google Color on the way out", () => {
    const { api, batchUpdateCalls } = recordingSheets();

    api.flush([
      {
        kind: "updateCell",
        sheetId: 1,
        rowIndex: 0,
        colIndex: 0,
        backgroundColor: lightGreen,
      },
    ]);

    expect(
      batchUpdateCalls[0]?.requests?.[0]?.updateCells?.rows?.[0]?.values?.[0]
        ?.userEnteredFormat?.backgroundColor,
    ).toEqual(lightGreen);
  });

  it("maps Google Color back to the RGB record", () => {
    const { api } = recordingSheets({
      sheets: [
        {
          properties: { sheetId: 1, title: "Records" },
          data: [
            {
              startColumn: 0,
              startRow: 0,
              columnMetadata: [{}],
              rowData: [
                {
                  values: [
                    {
                      effectiveValue: { stringValue: "x" },
                      userEnteredFormat: { backgroundColor: lightGreen },
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    });

    expect(
      api.fetchSheetProperties().sheets[0]?.gridBlocks?.[0]?.rows[0]?.cells[0]
        ?.backgroundColor,
    ).toEqual(lightGreen);
  });
});

describe("GoogleSheetsAPI HTTP transport", () => {
  function seedApi(props: { isDryRun?: boolean } = {}) {
    const transport = vi.fn((_request: SheetsHttpRequest) => ({
      spreadsheetId,
      sheets: [],
    }));
    const reported: GoogleRequest[] = [];
    const api = GoogleSheetsAPI.initHttp({
      spreadsheetId,
      transport,
      isDryRun: props.isDryRun ?? false,
      reportRequests: (requests) => reported.push(...requests),
    });
    return { api, transport, reported };
  }

  it("sends one GET for sheet properties, carrying the field mask", () => {
    const { api, transport } = seedApi();
    const fields =
      "properties(timeZone),sheets(properties(sheetId,title),tables(tableId,name,range))";

    api.fetchSheetProperties();

    expect(transport).toHaveBeenCalledWith({
      method: "GET",
      url:
        `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}` +
        `?fields=${encodeURIComponent(fields)}`,
      body: null,
    });
  });

  it("posts grid ranges as a getByDataFilter body", () => {
    const { api, transport } = seedApi();
    const gridRanges = [{ sheetId: 111, startRowIndex: 0 }];

    api.fetchGrid(gridRanges, {
      includeProgrammaticFacts: false,
    });

    expect(transport.mock.calls[0]?.[0].method).toBe("POST");
    expect(transport.mock.calls[0]?.[0].url).toContain(":getByDataFilter");
    expect(JSON.parse(transport.mock.calls[0]?.[0].body ?? "")).toEqual({
      dataFilters: [{ gridRange: gridRanges[0] }],
      includeGridData: true,
    });
  });

  it("posts mapped writes to batchUpdate when the run is not a dry run", () => {
    const { api, transport } = seedApi();

    api.flush([
      {
        kind: "updateCell",
        sheetId: 111,
        rowIndex: 5,
        colIndex: 2,
        value: "x",
      },
    ]);

    expect(transport.mock.calls[0]?.[0].url).toContain(":batchUpdate");
    expect(
      JSON.parse(transport.mock.calls[0]?.[0].body ?? "").requests,
    ).toHaveLength(1);
  });

  it("sends nothing at all on a dry run, and reports what it withheld", () => {
    const { api, transport, reported } = seedApi({ isDryRun: true });

    api.flush([
      {
        kind: "updateCell",
        sheetId: 111,
        rowIndex: 5,
        colIndex: 2,
        value: "x",
      },
    ]);

    expect(transport).not.toHaveBeenCalled();
    expect(reported).toHaveLength(1);
    expect(reported[0]?.updateCells).toBeDefined();
  });

  it("still reads from the live spreadsheet on a dry run", () => {
    const { api, transport } = seedApi({ isDryRun: true });

    api.fetchSheetProperties();
    api.fetchGrid([{ sheetId: 111, startRowIndex: 0 }], {
      includeProgrammaticFacts: false,
    });

    expect(transport).toHaveBeenCalledTimes(2);
  });
});

describe("GoogleSheetsAPI spreadsheet binding", () => {
  it("sends its bound spreadsheet id to get, getByDataFilter and batchUpdate", () => {
    const { api, requestedIds } = recordingSheets();

    api.fetchSheetProperties();
    api.fetchGrid([{ sheetId: 111, startRowIndex: 0 }], {
      includeProgrammaticFacts: false,
    });
    api.flush([
      {
        kind: "updateCell",
        sheetId: 111,
        rowIndex: 5,
        colIndex: 2,
        value: "x",
      },
    ]);

    expect(requestedIds).toEqual([spreadsheetId, spreadsheetId, spreadsheetId]);
  });

  it("asks for a bound script when Apps Script has no active spreadsheet", () => {
    stubScriptAndSpreadsheetApp({ spreadsheetId: null });
    vi.stubGlobal("Sheets", { Spreadsheets: {} });

    expect(() => GoogleSheetsAPI.forAppsScript()).toThrowError(
      "bind the Apps Script project to its spreadsheet",
    );
    vi.unstubAllGlobals();
  });
});

describe("GoogleSheetsAPI raw request", () => {
  it("sends a wrapped raw request out of flush unchanged", () => {
    const { api, batchUpdateCalls } = recordingSheets();
    const request: GoogleRequest = {
      updateDimensionProperties: {
        range: {
          sheetId: 111,
          dimension: "COLUMNS",
          startIndex: 0,
          endIndex: 1,
        },
        properties: { pixelSize: 40 },
        fields: "pixelSize",
      },
    };

    api.flush([{ kind: "raw", request: googleRawRequest(request) }]);

    expect(batchUpdateCalls).toEqual([{ requests: [request] }]);
  });

  it("will not take a Google request that was not wrapped", () => {
    const request: GoogleRequest = { updateTable: { table: { tableId: "t" } } };

    // @ts-expect-error A raw write takes only the port's opaque type.
    const operation: LocalWriteOperation = { kind: "raw", request };

    expect(operation.kind).toBe("raw");
  });
});
