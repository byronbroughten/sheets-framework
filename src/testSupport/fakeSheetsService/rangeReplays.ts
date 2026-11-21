import { Obj } from "@byronbroughten/utils/obj";

import type { CopyPasteType } from "../../00_Source/RawSource/RawSource.ts";
import type { FakeCell } from "../fakeSheetsService.ts";
import { a1Notation } from "./a1Notation.ts";
import { fakeCells } from "./fakeCells.ts";
import { type BandChange, type BoundedRange, fakeGrid } from "./fakeGrid.ts";
import {
  type FakeSheetState,
  type FakeSpreadsheet,
  fakeSpreadsheet,
  type FakeTableState,
} from "./fakeSpreadsheet.ts";
import { ruleFormulas } from "./ruleFormulas.ts";
import { unreplayed } from "./unreplayed.ts";

type Response = GoogleAppsScript.Sheets.Schema.Response;
type GridRange = GoogleAppsScript.Sheets.Schema.GridRange;
type ConditionalFormatRule =
  GoogleAppsScript.Sheets.Schema.ConditionalFormatRule;

type CellPaste = (cell: FakeCell, source: FakeCell) => FakeCell;

// Measured live (sheets-framework#61): a format paste carries no cell validation, and a validation paste no format.
const cellPastes: Record<CopyPasteType, CellPaste> = {
  PASTE_FORMAT(cell, source) {
    return fakeCells.withPastedFacts(cell, source, [
      "backgroundColor",
      "numberFormatType",
    ]);
  },
  PASTE_DATA_VALIDATION(cell, source) {
    return fakeCells.withPastedFacts(cell, source, [
      "dataValidationConditionType",
    ]);
  },
  PASTE_FORMULA: withPastedFormula,
};

// The fake keeps a formula's text, so it can't shift a relative reference as Sheets would.

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
    if (!Obj.isKey(cellPastes, pasteType)) {
      throw unreplayed(`copyPaste type ${pasteType}`);
    }
    if ((request.pasteOrientation ?? "NORMAL") !== "NORMAL") {
      throw unreplayed(`copyPaste orientation ${request.pasteOrientation}`);
    }
    const sheetId = request.source?.sheetId;
    if (request.destination?.sheetId !== sheetId) {
      throw unreplayed("copyPaste across sheets");
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
        cellPastes[pasteType](cell, sourceCell),
      );
    });
    if (pasteType === "PASTE_FORMAT") {
      extendConditionalFormats(sheet, { source, destination });
    }
    return {};
  },
  // Measured live (sheets-framework#204, #364): the whole cell moves, its Table column's name, type and rule with it, at the grid's edge too, and the source is left empty.
  cutPaste(
    spreadsheet: FakeSpreadsheet,
    request: GoogleAppsScript.Sheets.Schema.CutPasteRequest,
  ): Response {
    const pasteType = request.pasteType ?? "PASTE_NORMAL";
    if (pasteType !== "PASTE_NORMAL") {
      throw unreplayed(`cutPaste type ${pasteType}`);
    }
    const sheetId = request.source?.sheetId;
    const coordinate = request.destination ?? {};
    if (coordinate.sheetId !== sheetId) {
      throw unreplayed("cutPaste across sheets");
    }
    const sheet = fakeSpreadsheet.sheet(spreadsheet, sheetId);
    const source = fakeGrid.rangeInGrid(sheet, request.source ?? {});
    const rowOffset = (coordinate.rowIndex ?? 0) - source.startRowIndex;
    const colOffset = (coordinate.columnIndex ?? 0) - source.startColumnIndex;
    const destination = {
      startRowIndex: source.startRowIndex + rowOffset,
      endRowIndex: source.endRowIndex + rowOffset,
      startColumnIndex: source.startColumnIndex + colOffset,
      endColumnIndex: source.endColumnIndex + colOffset,
    };
    if (overlaps(source, destination)) {
      throw unreplayed("a cutPaste onto its own source");
    }
    sheet.conditionalFormats = cutRules(sheet, { source, destination });
    const lifted: FakeCell[][] = [];
    fakeGrid.forEachCell(source, (rowIndex, colIndex) => {
      const row = (lifted[rowIndex - source.startRowIndex] ??= []);
      row.push(fakeGrid.cell(sheet, rowIndex, colIndex));
      fakeGrid.setCell(sheet, rowIndex, colIndex, null);
    });
    fakeGrid.forEachCell(destination, (rowIndex, colIndex) => {
      fakeGrid.setCell(
        sheet,
        rowIndex,
        colIndex,
        lifted[rowIndex - destination.startRowIndex]?.[
          colIndex - destination.startColumnIndex
        ] ?? null,
      );
    });
    for (const table of tablesWithColumnsLifted(sheet, source)) {
      carryTableColumns({
        spreadsheet,
        sheet,
        table,
        move: { source, destination },
      });
    }
    return {};
  },
};

interface PasteRanges {
  source: BoundedRange;
  destination: BoundedRange;
}

interface TableColumnCut {
  spreadsheet: FakeSpreadsheet;
  sheet: FakeSheetState;
  table: FakeTableState;
  move: PasteRanges;
}

// Only a cut lifting a Table's whole columns was measured.
function tablesWithColumnsLifted(
  sheet: FakeSheetState,
  source: BoundedRange,
): FakeTableState[] {
  return sheet.tables.filter(
    (table) => coversRows(source, table) && coversColumns(table, source),
  );
}

// Measured live: a cut inside a Table carries its columns' facts; Scratch353D's paste out of the Table made a new one-column Table over it.
function carryTableColumns({
  spreadsheet,
  sheet,
  table,
  move,
}: TableColumnCut): void {
  const facts = tableColumnFacts(table, move);
  if (coversColumns(table, move.destination)) {
    Object.assign(table, facts.merged);
    return;
  }
  if (!isLastColumnToBefore(table, move)) {
    throw unreplayed(
      "a cutPaste of a Table's columns out of the Table, other than its last column to just before it",
    );
  }
  Object.assign(table, facts.kept);
  sheet.tables.push({
    tableId: `fake-table-${sheet.sheetId}-cut-${sheet.tables.length}`,
    name: newTableName(spreadsheet),
    startRowIndex: table.startRowIndex,
    endRowIndex: table.endRowIndex,
    startColumnIndex: move.destination.startColumnIndex,
    endColumnIndex: move.destination.endColumnIndex,
    ...facts.carried,
  });
}

type ColumnFacts = Pick<
  FakeTableState,
  "columnTypes" | "columnValidationValues" | "columnValidationConditionTypes"
>;

// The facts the cut leaves, those it carries, and both together where the paste lands inside the Table.
interface TableColumnFacts {
  kept: ColumnFacts;
  carried: ColumnFacts;
  merged: ColumnFacts;
}

function tableColumnFacts(
  table: FakeTableState,
  { source, destination }: PasteRanges,
): TableColumnFacts {
  const colOffset = destination.startColumnIndex - source.startColumnIndex;
  const types = splitColumnFacts(table.columnTypes, source, colOffset);
  const values = splitColumnFacts(
    table.columnValidationValues,
    source,
    colOffset,
  );
  const conditionTypes = splitColumnFacts(
    table.columnValidationConditionTypes,
    source,
    colOffset,
  );
  return {
    kept: {
      columnTypes: types?.kept,
      columnValidationValues: values?.kept,
      columnValidationConditionTypes: conditionTypes?.kept,
    },
    carried: {
      columnTypes: types?.carried,
      columnValidationValues: values?.carried,
      columnValidationConditionTypes: conditionTypes?.carried,
    },
    merged: {
      columnTypes: mergedColumnFacts(types),
      columnValidationValues: mergedColumnFacts(values),
      columnValidationConditionTypes: mergedColumnFacts(conditionTypes),
    },
  };
}

// Inferred: Scratch353D's one existing Table gave Table2, so n is one past the spreadsheet's Table count.
function newTableName(spreadsheet: FakeSpreadsheet): string {
  const name = `Table${tableCount(spreadsheet) + 1}`;
  const isTaken = spreadsheet.sheets.some((sheet) =>
    sheet.tables.some((table) => table.name === name),
  );
  if (isTaken) {
    throw unreplayed(
      `a cutPaste making a new Table when ${name} is already taken, a name no scratch has measured`,
    );
  }
  return name;
}

// A one-column cut of a span's last column onto the column just left of it, as Scratch353D did.
function isLastColumnToBefore(span: BoundedRange, move: PasteRanges): boolean {
  const { source, destination } = move;
  return (
    source.endColumnIndex - source.startColumnIndex === 1 &&
    source.endColumnIndex === span.endColumnIndex &&
    destination.endColumnIndex === span.startColumnIndex &&
    source.startRowIndex === destination.startRowIndex
  );
}

function tableCount(spreadsheet: FakeSpreadsheet): number {
  return spreadsheet.sheets.reduce(
    (count, sheet) => count + sheet.tables.length,
    0,
  );
}

interface SplitColumnFacts<FT> {
  kept: Record<number, FT>;
  carried: Record<number, FT>;
}

function splitColumnFacts<FT>(
  byColumn: Record<number, FT> | undefined,
  source: BoundedRange,
  colOffset: number,
): SplitColumnFacts<FT> | undefined {
  if (byColumn === undefined) return undefined;
  return Object.entries(byColumn).reduce<SplitColumnFacts<FT>>(
    (split, [key, fact]) => {
      const colIndex = Number(key);
      if (isInColumns(source, colIndex)) {
        split.carried[colIndex + colOffset] = fact;
      } else {
        split.kept[colIndex] = fact;
      }
      return split;
    },
    { kept: {}, carried: {} },
  );
}

// A carried fact replaces the one its destination column held.
function mergedColumnFacts<FT>(
  split: SplitColumnFacts<FT> | undefined,
): Record<number, FT> | undefined {
  if (split === undefined) return undefined;
  return { ...split.kept, ...split.carried };
}

const ruleCutRefusals = {
  cut: "a cutPaste under a conditional-format rule, other than a sideways cut inside a whole-Table rule naming no cut or pasted column, or naming one on a one-column cut to the right, the pasted column only $-anchored, or a whole-Table rule's last column cut to just before it, naming neither",
  outside:
    "a cutPaste under a conditional-format rule outside the cut source, over the destination or not, whose formula names a cut or pasted column, a shape not yet measured",
} as const;

function cutRules(
  sheet: FakeSheetState,
  move: PasteRanges,
): ConditionalFormatRule[] | undefined {
  return sheet.conditionalFormats?.map((rule) => {
    const isCut = (rule.ranges ?? []).some((range) =>
      overlaps(fakeGrid.boundedRange(sheet, range), move.source),
    );
    const isNamingMove = ruleColumnIndexes(rule).some(
      (colIndex) =>
        isInColumns(move.source, colIndex) ||
        isInColumns(move.destination, colIndex),
    );
    // Unmeasured: no scratch has cut a column a rule outside the cut source names, one over the destination included, so throw rather than guess.
    if (!isCut && isNamingMove) throw unreplayed(ruleCutRefusals.outside);
    if (!isCut) return rule;
    const range = wholeTableRange(sheet, rule);
    if (range === undefined) throw unreplayed(ruleCutRefusals.cut);
    const isLeavingTable =
      isLastColumnToBefore(range, move) && coversRows(move.source, range);
    if (isLeavingTable && !isNamingMove) return withRangeOverPaste(rule, range);
    if (!isSidewaysInRange(range, move)) throw unreplayed(ruleCutRefusals.cut);
    if (!isNamingMove) return rule;
    return withPasteReferencesShifted(rule, move);
  });
}

// The one range of a rule over a whole Table's columns.
function wholeTableRange(
  sheet: FakeSheetState,
  rule: ConditionalFormatRule,
): BoundedRange | undefined {
  const [range] = rule.ranges ?? [];
  if (range === undefined || (rule.ranges ?? []).length > 1) return undefined;
  const bounded = fakeGrid.boundedRange(sheet, range);
  const isOverWholeTable = sheet.tables.some(
    (table) =>
      table.startColumnIndex === bounded.startColumnIndex &&
      table.endColumnIndex === bounded.endColumnIndex,
  );
  if (!isOverWholeTable) return undefined;
  return bounded;
}

// Measured live (sheets-framework#217, Scratch217A; #353, Scratch353A to C): only this shape was seen.
function isSidewaysInRange(range: BoundedRange, move: PasteRanges): boolean {
  const { source, destination } = move;
  if (source.startRowIndex !== destination.startRowIndex) return false;
  return coversColumns(range, source) && coversColumns(range, destination);
}

// Inferred from Scratch353D's whole batch: the rule moves one column left, over the paste and off the emptied column, its formula as written.
function withRangeOverPaste(
  rule: ConditionalFormatRule,
  range: BoundedRange,
): ConditionalFormatRule {
  const [gridRange] = rule.ranges ?? [];
  return {
    ...rule,
    ranges: [
      {
        ...gridRange,
        startColumnIndex: range.startColumnIndex - 1,
        endColumnIndex: range.endColumnIndex - 1,
      },
    ],
  };
}

// Measured (Scratch353A to C): a cut column's references stay as written; one into the paste takes the band insert's shift, which the fake's insert leaves out.
function withPasteReferencesShifted(
  rule: ConditionalFormatRule,
  { source, destination }: PasteRanges,
): ConditionalFormatRule {
  const width = destination.endColumnIndex - destination.startColumnIndex;
  const isOneColumnRight =
    width === 1 && source.endColumnIndex <= destination.startColumnIndex;
  if (!isOneColumnRight) throw unreplayed(ruleCutRefusals.cut);
  return ruleFormulas.mapped(rule, (formula) =>
    a1Notation.mapCellReferences(
      formula,
      ({ reference, anchor, letters, row }) => {
        const colIndex = a1Notation.columnIndex(letters);
        if (!isInColumns(destination, colIndex)) return reference;
        if (anchor !== "$") throw unreplayed(ruleCutRefusals.cut);
        return `$${a1Notation.columnLetters(colIndex + width)}${row}`;
      },
    ),
  );
}

function ruleColumnIndexes(rule: ConditionalFormatRule): number[] {
  const values = rule.booleanRule?.condition?.values ?? [];
  return values.flatMap((value) =>
    a1Notation
      .cellReferencesIn(value.userEnteredValue ?? "")
      .map(({ letters }) => a1Notation.columnIndex(letters)),
  );
}

function isInColumns(range: BoundedRange, colIndex: number): boolean {
  return range.startColumnIndex <= colIndex && colIndex < range.endColumnIndex;
}

// Refuses a plain value or a cell reference rather than guess what Sheets would paste.
function withPastedFormula(cell: FakeCell, source: FakeCell): FakeCell {
  const formula = fakeCells.value(source);
  const isFormula = typeof source === "object" && source?.isFormula === true;
  if (!isFormula || typeof formula !== "string") {
    throw unreplayed("a formula paste from a cell holding no formula");
  }
  if (a1Notation.cellReference.test(formula)) {
    throw unreplayed(
      `a formula paste of ${formula}, which holds a cell reference`,
    );
  }
  return fakeCells.withValue(cell, formula, true);
}

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
  paste: PasteRanges,
): void {
  sheet.conditionalFormats = sheet.conditionalFormats?.map((rule) => ({
    ...rule,
    ranges: (rule.ranges ?? []).map((range) => extendedRuleRange(range, paste)),
  }));
}

function extendedRuleRange(
  range: GridRange,
  { source, destination }: PasteRanges,
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
    throw unreplayed(
      "a format paste onto rows other than those just below a rule's range",
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

function coversColumns(outer: BoundedRange, inner: BoundedRange): boolean {
  return (
    outer.startColumnIndex <= inner.startColumnIndex &&
    inner.endColumnIndex <= outer.endColumnIndex
  );
}

function coversRows(outer: BoundedRange, inner: BoundedRange): boolean {
  return (
    outer.startRowIndex <= inner.startRowIndex &&
    inner.endRowIndex <= outer.endRowIndex
  );
}
