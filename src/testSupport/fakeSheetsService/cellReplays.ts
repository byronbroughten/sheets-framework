import type { FakeCellValue } from "../fakeSheetsService";
import { fakeCells } from "./fakeCells";
import { type BoundedRange, fakeGrid } from "./fakeGrid";
import {
  type FakeSheetState,
  type FakeSpreadsheet,
  fakeSpreadsheet,
} from "./fakeSpreadsheet";

type Response = GoogleAppsScript.Sheets.Schema.Response;
type RowData = GoogleAppsScript.Sheets.Schema.RowData;

export const cellReplays = {
  // With a range, a cell the rows leave out is cleared of the masked fields, as live.
  updateCells(
    spreadsheet: FakeSpreadsheet,
    request: GoogleAppsScript.Sheets.Schema.UpdateCellsRequest,
  ): Response {
    const fields = fakeCells.fields(request.fields, "updateCells");
    const rows = request.rows ?? [];
    const sheetId = request.range?.sheetId ?? request.start?.sheetId;
    const sheet = fakeSpreadsheet.sheet(spreadsheet, sheetId);
    const range = updateCellsRange(request, rows);
    fakeGrid.forEachCell(range, (rowIndex, colIndex) => {
      const data =
        rows[rowIndex - range.startRowIndex]?.values?.[
          colIndex - range.startColumnIndex
        ] ?? {};
      fakeGrid.updateCell(sheet, rowIndex, colIndex, (cell) =>
        fakeCells.withCellData(cell, data, fields),
      );
    });
    return {};
  },
  repeatCell(
    spreadsheet: FakeSpreadsheet,
    request: GoogleAppsScript.Sheets.Schema.RepeatCellRequest,
  ): Response {
    const fields = fakeCells.fields(request.fields, "repeatCell");
    const sheet = fakeSpreadsheet.sheet(spreadsheet, request.range?.sheetId);
    const range = fakeGrid.boundedRange(sheet, request.range ?? {});
    fakeGrid.forEachCell(range, (rowIndex, colIndex) => {
      fakeGrid.updateCell(sheet, rowIndex, colIndex, (cell) =>
        fakeCells.withCellData(cell, request.cell ?? {}, fields),
      );
    });
    return {};
  },
  // Formulas read back as their text; the fake evaluates nothing.
  pasteData(
    spreadsheet: FakeSpreadsheet,
    request: GoogleAppsScript.Sheets.Schema.PasteDataRequest,
  ): Response {
    const type = request.type ?? "PASTE_NORMAL";
    if (type !== "PASTE_NORMAL" && type !== "PASTE_FORMULA") {
      throw new Error(
        `The fake Sheets service does not replay pasteData type ${type}.`,
      );
    }
    if (request.html === true) {
      throw new Error(
        "The fake Sheets service does not replay HTML pasteData.",
      );
    }
    const { coordinate } = request;
    const sheet = fakeSpreadsheet.sheet(spreadsheet, coordinate?.sheetId);
    const startRow = coordinate?.rowIndex ?? 0;
    const startColumn = coordinate?.columnIndex ?? 0;
    parseDelimited(request.data ?? "", request.delimiter ?? ",").forEach(
      (record, rowOffset) => {
        record.forEach((field, colOffset) => {
          const rowIndex = startRow + rowOffset;
          const colIndex = startColumn + colOffset;
          fakeGrid.updateCell(sheet, rowIndex, colIndex, (cell) =>
            fakeCells.withValue(
              cell,
              pastedValue(field),
              field.startsWith("="),
            ),
          );
        });
      },
    );
    return {};
  },
  setDataValidation(
    spreadsheet: FakeSpreadsheet,
    request: GoogleAppsScript.Sheets.Schema.SetDataValidationRequest,
  ): Response {
    const sheet = fakeSpreadsheet.sheet(spreadsheet, request.range?.sheetId);
    const range = fakeGrid.boundedRange(sheet, request.range ?? {});
    const fields = fakeCells.fields("dataValidation", "setDataValidation");
    fakeGrid.forEachCell(range, (rowIndex, colIndex) => {
      fakeGrid.updateCell(sheet, rowIndex, colIndex, (cell) =>
        fakeCells.withCellData(
          cell,
          request.rule === undefined ? {} : { dataValidation: request.rule },
          fields,
        ),
      );
    });
    return {};
  },
  // Text cells only, and formulas only when asked; a number never matches.
  findReplace(
    spreadsheet: FakeSpreadsheet,
    request: GoogleAppsScript.Sheets.Schema.FindReplaceRequest,
  ): Response {
    const pattern = findPattern(request);
    const replacement = request.replacement ?? "";
    let occurrencesChanged = 0;
    let valuesChanged = 0;
    findReplaceScopes(spreadsheet, request).forEach(({ sheet, range }) => {
      fakeGrid.forEachCell(range, (rowIndex, colIndex) => {
        const cell = fakeGrid.cell(sheet, rowIndex, colIndex);
        const value = fakeCells.value(cell);
        const isFormula = typeof cell === "object" && cell?.isFormula === true;
        if (typeof value !== "string") return;
        if (isFormula && request.includeFormulas !== true) return;
        const matchCount = value.match(pattern)?.length ?? 0;
        if (matchCount === 0) return;
        const replaced = request.searchByRegex
          ? value.replace(pattern, replacement)
          : value.replace(pattern, () => replacement);
        occurrencesChanged += matchCount;
        valuesChanged += 1;
        fakeGrid.setCell(
          sheet,
          rowIndex,
          colIndex,
          fakeCells.withValue(cell, replaced, isFormula),
        );
      });
    });
    return { findReplace: { occurrencesChanged, valuesChanged } };
  },
};

function updateCellsRange(
  request: GoogleAppsScript.Sheets.Schema.UpdateCellsRequest,
  rows: RowData[],
): BoundedRange {
  const widestRow = Math.max(0, ...rows.map((row) => row.values?.length ?? 0));
  const startRowIndex =
    request.range?.startRowIndex ?? request.start?.rowIndex ?? 0;
  const startColumnIndex =
    request.range?.startColumnIndex ?? request.start?.columnIndex ?? 0;
  return {
    startRowIndex,
    endRowIndex: request.range?.endRowIndex ?? startRowIndex + rows.length,
    startColumnIndex,
    endColumnIndex:
      request.range?.endColumnIndex ?? startColumnIndex + widestRow,
  };
}

// RFC 4180 records: a quoted field may hold the delimiter, a newline or a doubled quote.
function parseDelimited(data: string, delimiter: string): string[][] {
  const records: string[][] = [];
  let record: string[] = [];
  let field = "";
  let isQuoted = false;
  for (let i = 0; i < data.length; i++) {
    const char = data[i];
    if (isQuoted) {
      if (char === '"' && data[i + 1] === '"') {
        field += '"';
        i++;
      } else if (char === '"') {
        isQuoted = false;
      } else {
        field += char;
      }
      continue;
    }
    if (char === '"') {
      isQuoted = true;
    } else if (char === delimiter) {
      record.push(field);
      field = "";
    } else if (char === "\n") {
      record.push(field);
      records.push(record);
      record = [];
      field = "";
    } else {
      field += char;
    }
  }
  record.push(field);
  records.push(record);
  return records;
}

function pastedValue(field: string): FakeCellValue {
  if (field === "") return null;
  if (
    !field.startsWith("=") &&
    field.trim() !== "" &&
    !Number.isNaN(Number(field))
  ) {
    return Number(field);
  }
  return field;
}

function findPattern(
  request: GoogleAppsScript.Sheets.Schema.FindReplaceRequest,
): RegExp {
  const find = request.find ?? "";
  const source = request.searchByRegex
    ? find
    : find.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const anchored = request.matchEntireCell ? `^(?:${source})$` : source;
  return new RegExp(anchored, request.matchCase ? "g" : "gi");
}

function findReplaceScopes(
  spreadsheet: FakeSpreadsheet,
  request: GoogleAppsScript.Sheets.Schema.FindReplaceRequest,
): { sheet: FakeSheetState; range: BoundedRange }[] {
  if (request.allSheets === true) {
    return spreadsheet.sheets.map((sheet) => ({
      sheet,
      range: fakeGrid.boundedRange(sheet, {}),
    }));
  }
  if (request.sheetId !== undefined) {
    const sheet = fakeSpreadsheet.sheet(spreadsheet, request.sheetId);
    return [{ sheet, range: fakeGrid.boundedRange(sheet, {}) }];
  }
  const sheet = fakeSpreadsheet.sheet(spreadsheet, request.range?.sheetId);
  return [{ sheet, range: fakeGrid.boundedRange(sheet, request.range ?? {}) }];
}
