import { describe, expect, it } from "vitest";

import type { GoogleRequest } from "../00_Source/GoogleSheets/GoogleSheetsAPI";
import { getSheetTraitByName } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import { UpdateRequestSummary } from "./UpdateRequestSummary";

const itemGid = getSheetTraitByName("item", "sheetGid");
const unknownGid = 999999;

function onlyLine(request: GoogleRequest): string {
  const [line] = UpdateRequestSummary.init([request]).lines;
  return (line ?? "").replace(/\s+/g, " ").trim();
}

describe("UpdateRequestSummary.lines", () => {
  it("names the sheet, the range, the values and the field mask of a cell write", () => {
    expect(
      onlyLine({
        updateCells: {
          range: {
            sheetId: itemGid,
            startRowIndex: 4,
            endRowIndex: 6,
            startColumnIndex: 2,
            endColumnIndex: 3,
          },
          rows: [
            { values: [{ userEnteredValue: { stringValue: "yes" } }] },
            { values: [{ userEnteredValue: { numberValue: 42 } }] },
          ],
          fields: "userEnteredValue",
        },
      }),
    ).toBe('updateCells item!C5:C6 2 cell(s) "yes", 42 [userEnteredValue]');
  });

  it("falls back to the gid for a sheet the config does not describe", () => {
    expect(
      onlyLine({
        updateCells: {
          range: { sheetId: unknownGid, startRowIndex: 0, endRowIndex: 1 },
        },
      }),
    ).toContain(`gid ${unknownGid}`);
  });

  it("counts every cell a column fill covers", () => {
    expect(
      onlyLine({
        repeatCell: {
          range: {
            sheetId: itemGid,
            startRowIndex: 4,
            endRowIndex: 9,
            startColumnIndex: 5,
            endColumnIndex: 6,
          },
          cell: { userEnteredValue: { boolValue: false } },
          fields: "userEnteredValue",
        },
      }),
    ).toBe("repeatCell item!F5:F9 5 cell(s) false [userEnteredValue]");
  });

  it("names the range, cell count and condition type of a data-validation rule", () => {
    expect(
      onlyLine({
        setDataValidation: {
          range: {
            sheetId: itemGid,
            startRowIndex: 4,
            endRowIndex: 5,
            startColumnIndex: 5,
            endColumnIndex: 6,
          },
          rule: { condition: { type: "BOOLEAN" } },
        },
      }),
    ).toBe("setDataValidation item!F5:F5 1 cell(s) BOOLEAN");
  });

  it("counts the rows an append adds", () => {
    expect(
      onlyLine({
        appendCells: {
          sheetId: itemGid,
          rows: [{ values: [{ userEnteredValue: { stringValue: "r:abc" } }] }],
          fields: "userEnteredValue",
        },
      }),
    ).toBe('appendCells item 1 row(s) "r:abc" [userEnteredValue]');
  });

  it("states a row delete as the 1-based rows it removes", () => {
    expect(
      onlyLine({
        deleteDimension: {
          range: {
            sheetId: itemGid,
            dimension: "ROWS",
            startIndex: 6,
            endIndex: 9,
          },
        },
      }),
    ).toBe("deleteDimension item!7:9 delete 3 rows");
  });

  it("states a column insert by its column letter", () => {
    expect(
      onlyLine({
        insertDimension: {
          range: {
            sheetId: itemGid,
            dimension: "COLUMNS",
            startIndex: 27,
            endIndex: 28,
          },
        },
      }),
    ).toBe("insertDimension item!AB:AB insert 1 columns");
  });

  it("states a sort by its column and direction", () => {
    expect(
      onlyLine({
        sortRange: {
          range: { sheetId: itemGid, startRowIndex: 4 },
          sortSpecs: [{ dimensionIndex: 2, sortOrder: "ASCENDING" }],
        },
      }),
    ).toBe("sortRange item!A5: by column C ascending");
  });

  it("states a range-scoped replace as its range, its two strings and its flags", () => {
    expect(
      onlyLine({
        findReplace: {
          find: "Old",
          replacement: "New",
          matchEntireCell: true,
          range: {
            sheetId: itemGid,
            startRowIndex: 4,
            endRowIndex: 9,
            startColumnIndex: 2,
            endColumnIndex: 3,
          },
        },
      }),
    ).toBe(
      'findReplace item!C5:C9 matching cells "Old" → "New" matchEntireCell',
    );
  });

  it("names the whole sheet when the replace is scoped by sheet rather than range", () => {
    expect(
      onlyLine({
        findReplace: {
          find: "Old",
          replacement: "New",
          sheetId: itemGid,
        },
      }),
    ).toBe('findReplace item!all matching cells "Old" → "New" (no flags)');
  });

  it("says every sheet when the replace is unscoped", () => {
    expect(
      onlyLine({
        findReplace: {
          find: "Old",
          replacement: "New",
          allSheets: true,
          includeFormulas: true,
        },
      }),
    ).toBe(
      'findReplace every sheet matching cells "Old" → "New" includeFormulas',
    );
  });

  it("names the sheet, coordinate, formula text and PASTE_FORMULA of a formula paste", () => {
    expect(
      onlyLine({
        pasteData: {
          coordinate: {
            sheetId: itemGid,
            rowIndex: 4,
            columnIndex: 6,
          },
          data: '"=2+SINGLE(item[Required count])"',
          delimiter: "\t",
          type: "PASTE_FORMULA",
        },
      }),
    ).toBe(
      "pasteData item!G5:G5 1 row(s) =2+SINGLE(item[Required count]) PASTE_FORMULA",
    );
  });

  it("spans the quoted records of a column formula fill", () => {
    expect(
      onlyLine({
        pasteData: {
          coordinate: {
            sheetId: itemGid,
            rowIndex: 4,
            columnIndex: 6,
          },
          data: '"=2+1"\n"=2+1"\n"=2+1"',
          delimiter: "\t",
          type: "PASTE_FORMULA",
        },
      }),
    ).toBe("pasteData item!G5:G7 3 row(s) =2+1 PASTE_FORMULA");
  });

  it("names the sheet, range and condition of an added conditional format rule", () => {
    expect(
      onlyLine({
        addConditionalFormatRule: {
          index: 0,
          rule: {
            ranges: [
              {
                sheetId: itemGid,
                startRowIndex: 4,
                endRowIndex: 11,
                startColumnIndex: 0,
                endColumnIndex: 1,
              },
            ],
            booleanRule: {
              condition: { type: "CUSTOM_FORMULA" },
            },
          },
        },
      }),
    ).toBe("addConditionalFormatRule item!A5:A11 prepend CUSTOM_FORMULA");
  });

  it("names the sheet and index of a deleted conditional format rule", () => {
    expect(
      onlyLine({
        deleteConditionalFormatRule: {
          sheetId: itemGid,
          index: 3,
        },
      }),
    ).toBe("deleteConditionalFormatRule item index 3");
  });

  it("names the range and kind of an added protected range", () => {
    expect(
      onlyLine({
        addProtectedRange: {
          protectedRange: {
            range: {
              sheetId: itemGid,
              startRowIndex: 4,
              endRowIndex: 11,
              startColumnIndex: 0,
              endColumnIndex: 1,
            },
            warningOnly: true,
            description: "id warning",
          },
        },
      }),
    ).toBe("addProtectedRange item!A5:A11 warning id warning");
  });

  it("names the id of a deleted protected range", () => {
    expect(
      onlyLine({
        deleteProtectedRange: { protectedRangeId: 11 },
      }),
    ).toBe("deleteProtectedRange (no sheet) id 11");
  });

  it("lists every column of a table column properties update in index order as Name: TYPE", () => {
    expect(
      onlyLine({
        updateTable: {
          table: {
            tableId: "t1",
            columnProperties: [
              { columnIndex: 1, columnName: "ID" },
              { columnName: "Name", columnType: "TEXT" },
            ],
          },
          fields: "columnProperties",
        },
      }),
    ).toBe("updateTable t1 2 cols Name: TEXT, ID: — [columnProperties]");
  });

  it("names a Table rename", () => {
    expect(
      onlyLine({
        updateTable: {
          table: { tableId: "t1", name: "spreadsheetConfig" },
          fields: "name",
        },
      }),
    ).toBe("updateTable t1 name spreadsheetConfig [name]");
  });

  it("names a tab title restore", () => {
    expect(
      onlyLine({
        updateSheetProperties: {
          properties: {
            sheetId: itemGid,
            title: "Spreadsheet Config",
          },
          fields: "title",
        },
      }),
    ).toBe("updateSheetProperties item title Spreadsheet Config [title]");
  });

  it("names the GID as a number, the title and the grid size of an added tab", () => {
    expect(
      onlyLine({
        addSheet: {
          properties: {
            sheetId: itemGid,
            title: "Spreadsheet Config",
            gridProperties: { rowCount: 20, columnCount: 6 },
          },
        },
      }),
    ).toBe(
      `addSheet gid ${itemGid} add tab Spreadsheet Config (20 rows × 6 cols)`,
    );
  });

  it("names the GID as a number, the range, the name and every column of an added Table", () => {
    expect(
      onlyLine({
        addTable: {
          table: {
            name: "spreadsheetConfig",
            range: {
              sheetId: itemGid,
              startRowIndex: 2,
              endRowIndex: 5,
              startColumnIndex: 1,
              endColumnIndex: 3,
            },
            columnProperties: [
              { columnIndex: 2, columnName: "Amount", columnType: "CURRENCY" },
              { columnIndex: 1, columnName: "Name", columnType: "TEXT" },
            ],
          },
        },
      }),
    ).toBe(
      `addTable gid ${itemGid}!B3:C5 add 2 cols spreadsheetConfig (Name: TEXT, Amount: CURRENCY)`,
    );
  });

  it("renders a request the framework does not model as its own verb and JSON", () => {
    const line = onlyLine({ deleteSheet: { sheetId: unknownGid } });

    expect(line).toContain("deleteSheet");
    expect(line).toContain(`gid ${unknownGid}`);
    expect(line).toContain(`"sheetId":${unknownGid}`);
  });

  it("is empty when nothing was queued", () => {
    const summary = UpdateRequestSummary.init([]);

    expect(summary.isEmpty).toBe(true);
    expect(summary.lines).toEqual([]);
  });
});
