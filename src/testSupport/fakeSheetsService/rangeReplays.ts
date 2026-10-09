import { Obj } from "@byronbroughten/utils/obj";

import type { CopyPasteType } from "../../00_Source/RawSource/RawSource";
import { fakeCells, type PastedFact } from "./fakeCells";
import { type BandChange, type BoundedRange, fakeGrid } from "./fakeGrid";
import {
  type FakeSheetState,
  type FakeSpreadsheet,
  fakeSpreadsheet,
} from "./fakeSpreadsheet";

type Response = GoogleAppsScript.Sheets.Schema.Response;
type GridRange = GoogleAppsScript.Sheets.Schema.GridRange;

// Measured live (sheets-framework#61): a format paste carries no cell validation, and a validation paste no format.
const pastedFacts: Record<CopyPasteType, readonly PastedFact[]> = {
  PASTE_FORMAT: ["backgroundColor", "numberFormatType"],
  PASTE_DATA_VALIDATION: ["dataValidationConditionType"],
};

export const rangeReplays = {
  insertRange(
    spreadsheet: FakeSpreadsheet,
    request: GoogleAppsScript.Sheets.Schema.InsertRangeRequest,
  ): Response {
    const sheet = fakeSpreadsheet.sheet(spreadsheet, request.range?.sheetId);
    fakeGrid.insertBand(
      sheet,
      bandChange(sheet, request.range, request.shiftDimension),
    );
    return {};
  },
  deleteRange(
    spreadsheet: FakeSpreadsheet,
    request: GoogleAppsScript.Sheets.Schema.DeleteRangeRequest,
  ): Response {
    const sheet = fakeSpreadsheet.sheet(spreadsheet, request.range?.sheetId);
    fakeGrid.removeBand(
      sheet,
      bandChange(sheet, request.range, request.shiftDimension),
    );
    return {};
  },
  // Measured live: one source row tiles over every destination row.
  copyPaste(
    spreadsheet: FakeSpreadsheet,
    request: GoogleAppsScript.Sheets.Schema.CopyPasteRequest,
  ): Response {
    const pasteType = request.pasteType ?? "PASTE_NORMAL";
    if (!Obj.isKey(pastedFacts, pasteType)) {
      throw new Error(
        `The fake Sheets service does not replay copyPaste type ${pasteType}.`,
      );
    }
    if ((request.pasteOrientation ?? "NORMAL") !== "NORMAL") {
      throw new Error(
        `The fake Sheets service does not replay copyPaste orientation ${request.pasteOrientation}.`,
      );
    }
    const facts = pastedFacts[pasteType];
    const sheetId = request.source?.sheetId;
    if (request.destination?.sheetId !== sheetId) {
      throw new Error(
        "The fake Sheets service does not replay copyPaste across sheets.",
      );
    }
    const sheet = fakeSpreadsheet.sheet(spreadsheet, sheetId);
    const source = fakeGrid.rangeInGrid(sheet, request.source ?? {});
    const destination = fakeGrid.rangeInGrid(sheet, request.destination ?? {});
    const sourceHeight = source.endRowIndex - source.startRowIndex;
    const sourceWidth = source.endColumnIndex - source.startColumnIndex;
    fakeGrid.forEachCell(destination, (rowIndex, colIndex) => {
      const sourceCell = fakeGrid.cell(
        sheet,
        source.startRowIndex +
          ((rowIndex - destination.startRowIndex) % sourceHeight),
        source.startColumnIndex +
          ((colIndex - destination.startColumnIndex) % sourceWidth),
      );
      fakeGrid.updateCell(sheet, rowIndex, colIndex, (cell) =>
        fakeCells.withPastedFacts(cell, sourceCell, facts),
      );
    });
    if (pasteType === "PASTE_FORMAT") {
      extendConditionalFormats(sheet, source, destination);
    }
    return {};
  },
};

function bandChange(
  sheet: FakeSheetState,
  range: GridRange | undefined,
  shiftDimension: string | undefined,
): BandChange {
  const bounded = fakeGrid.rangeInGrid(sheet, range ?? {});
  if (shiftDimension === "ROWS") {
    return validBand({
      dimension: "ROWS",
      startIndex: bounded.startRowIndex,
      count: bounded.endRowIndex - bounded.startRowIndex,
      crossStartIndex: bounded.startColumnIndex,
      crossEndIndex: bounded.endColumnIndex,
    });
  }
  if (shiftDimension === "COLUMNS") {
    return validBand({
      dimension: "COLUMNS",
      startIndex: bounded.startColumnIndex,
      count: bounded.endColumnIndex - bounded.startColumnIndex,
      crossStartIndex: bounded.startRowIndex,
      crossEndIndex: bounded.endRowIndex,
    });
  }
  throw new Error(
    `A range shift needs ROWS or COLUMNS, not ${shiftDimension}.`,
  );
}

function validBand(change: BandChange): BandChange {
  if (change.count <= 0 || change.crossEndIndex <= change.crossStartIndex) {
    throw new Error(
      `A range shift needs a non-empty range, not ${JSON.stringify(change)}.`,
    );
  }
  return change;
}

// Measured live (sheets-framework#61): a format paste onto the rows just below a rule's range extends the rule.
function extendConditionalFormats(
  sheet: FakeSheetState,
  source: BoundedRange,
  destination: BoundedRange,
): void {
  sheet.conditionalFormats = sheet.conditionalFormats?.map((rule) => ({
    ...rule,
    ranges: (rule.ranges ?? []).map((range) =>
      extendedRuleRange(range, source, destination),
    ),
  }));
}

function extendedRuleRange(
  range: GridRange,
  source: BoundedRange,
  destination: BoundedRange,
): GridRange {
  const bounded = {
    startRowIndex: range.startRowIndex ?? 0,
    endRowIndex: range.endRowIndex ?? Number.MAX_SAFE_INTEGER,
    startColumnIndex: range.startColumnIndex ?? 0,
    endColumnIndex: range.endColumnIndex ?? Number.MAX_SAFE_INTEGER,
  };
  if (!overlaps(bounded, source) || coversRows(bounded, destination)) {
    return range;
  }
  const isExtendable =
    bounded.startRowIndex <= source.startRowIndex &&
    bounded.endRowIndex === destination.startRowIndex &&
    source.endRowIndex === destination.startRowIndex &&
    source.startColumnIndex <= bounded.startColumnIndex &&
    bounded.endColumnIndex <= source.endColumnIndex &&
    source.startColumnIndex === destination.startColumnIndex &&
    source.endColumnIndex === destination.endColumnIndex;
  if (!isExtendable) {
    throw new Error(
      "The fake Sheets service does not replay a format paste onto rows other than those just below a rule's range.",
    );
  }
  return { ...range, endRowIndex: destination.endRowIndex };
}

function overlaps(a: BoundedRange, b: BoundedRange): boolean {
  return (
    a.startRowIndex < b.endRowIndex &&
    b.startRowIndex < a.endRowIndex &&
    a.startColumnIndex < b.endColumnIndex &&
    b.startColumnIndex < a.endColumnIndex
  );
}

function coversRows(outer: BoundedRange, inner: BoundedRange): boolean {
  return (
    outer.startRowIndex <= inner.startRowIndex &&
    inner.endRowIndex <= outer.endRowIndex
  );
}
