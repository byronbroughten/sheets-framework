import type {
  GoogleRequest,
  ModeledRequestVerb,
} from "../../00_Source/GoogleSheets/GoogleSheetsAPI";
import type {
  CellAfter,
  GoogleCellData,
  GoogleGridRange,
  GridCell,
} from "./flushClassifierTypes";
import type {
  BoundedRange,
  DimensionShift,
  ShiftDimension,
} from "./FlushTracker";
import {
  boundedCells,
  type CellWrite,
  type RequestJudge,
} from "./RequestJudge";

type DimensionRange = GoogleAppsScript.Sheets.Schema.DimensionRange;

// cutPaste and moveDimension aren't built yet, but #219's column move sends cutPaste, so both are judged ahead of their builders.
export type ClassifiedVerb = ModeledRequestVerb | "cutPaste" | "moveDimension";

type RequestJudgement<CV extends ClassifiedVerb> = (
  inner: NonNullable<GoogleRequest[CV]>,
  judge: RequestJudge,
) => void;

// Pastes that carry values; the rest carry formats, validation or rules only.
const valuePasteTypes: ReadonlySet<string> = new Set([
  "PASTE_NORMAL",
  "PASTE_VALUES",
  "PASTE_FORMULA",
  "PASTE_NO_BORDERS",
]);

// A request kind with no entry here is held as unclassified, so naming a verb in ModeledRequestVerb obliges an entry.
export const requestJudgements: {
  [CV in ClassifiedVerb]: RequestJudgement<CV>;
} = {
  appendCells(): void {},
  appendDimension(append, judge): void {
    judge.tracker.grow(
      append.sheetId ?? 0,
      shiftDimension(append.dimension),
      append.length ?? 0,
    );
  },
  insertDimension(insert, judge): void {
    const range = dimensionSpan(insert.range, judge);
    if (range === undefined) return;
    judge.tracker.shift({ ...range, band: undefined });
  },
  deleteDimension(remove, judge): void {
    const range = dimensionSpan(remove.range, judge);
    if (range === undefined) return;
    judge.delete(judge.cellsIn(dimensionGridRange(range)));
    judge.tracker.shift({ ...range, count: -range.count, band: undefined });
  },
  insertRange(insert, judge): void {
    const measured = rangeShiftOf(insert, judge);
    if (measured === undefined) return;
    judge.tracker.shift(measured.shift);
  },
  deleteRange(remove, judge): void {
    const measured = rangeShiftOf(remove, judge);
    if (measured === undefined) return;
    judge.delete(boundedCells(measured.bounded));
    judge.tracker.shift({ ...measured.shift, count: -measured.shift.count });
  },
  copyPaste(copy, judge): void {
    if (!valuePasteTypes.has(copy.pasteType ?? "PASTE_NORMAL")) return;
    if (!copy.source || !copy.destination) {
      judge.holdUnmeasured("a copy without a source or destination");
      return;
    }
    const source = judge.bounded(copy.source);
    const destination = judge.bounded(copy.destination);
    if (!source || !destination) return;
    const sourceHeight = source.rows.endIndex - source.rows.startIndex;
    const sourceWidth = source.cols.endIndex - source.cols.startIndex;
    const pasted = {
      sheetGid: destination.sheetGid,
      rows: tiledSpan(destination.rows, sourceHeight),
      cols: tiledSpan(destination.cols, sourceWidth),
    };
    const afters = boundedCells(pasted).map((cell) => ({
      cell,
      after: judge.afterFrom({
        sheetGid: source.sheetGid,
        rowIndex:
          source.rows.startIndex +
          ((cell.rowIndex - pasted.rows.startIndex) % sourceHeight),
        colIndex:
          source.cols.startIndex +
          ((cell.colIndex - pasted.cols.startIndex) % sourceWidth),
      }),
    }));
    judge.write(afters);
  },
  cutPaste(cut, judge): void {
    const coordinate = cut.destination;
    if (!cut.source || !coordinate) {
      judge.holdUnmeasured("a cut without a source or destination");
      return;
    }
    const source = judge.bounded(cut.source);
    if (!source) return;
    const destination: GridCell = {
      sheetGid: coordinate.sheetId ?? 0,
      rowIndex: coordinate.rowIndex ?? 0,
      colIndex: coordinate.columnIndex ?? 0,
    };
    // A cell the cut lands on that it also lifts from loses nothing: its content moves with the rest.
    const overwritten = boundedCells(source)
      .map((from) => ({
        from,
        cell: {
          sheetGid: destination.sheetGid,
          rowIndex:
            destination.rowIndex + from.rowIndex - source.rows.startIndex,
          colIndex:
            destination.colIndex + from.colIndex - source.cols.startIndex,
        },
      }))
      .filter(({ cell }) => !isInBounded(source, cell));
    judge.write(
      overwritten.map(({ from, cell }) => ({
        cell,
        after: judge.afterFrom(from),
      })),
    );
    judge.tracker.cut(source, destination);
  },
  moveDimension(move, judge): void {
    const range = dimensionSpan(move.source, judge);
    if (range === undefined) return;
    if (move.destinationIndex === undefined) {
      judge.holdUnmeasured("a move without a destination");
      return;
    }
    judge.tracker.move({
      sheetGid: range.sheetGid,
      dimension: range.dimension,
      startIndex: range.startIndex,
      endIndex: range.startIndex + range.count,
      destinationIndex: move.destinationIndex,
    });
  },
  findReplace(findReplace, judge): void {
    judge.holdFindReplace(
      findReplace.find ?? "",
      findReplace.replacement ?? "",
    );
  },
  sortRange(sort, judge): void {
    if (sort.range) judge.tracker.reorder(sort.range);
  },
  addConditionalFormatRule(): void {},
  deleteConditionalFormatRule(): void {},
  addProtectedRange(): void {},
  deleteProtectedRange(): void {},
  addSheet(add, judge): void {
    const properties = add.properties;
    // A tab Google numbers itself can't be named by a later request in the same flush.
    if (properties?.sheetId === undefined) return;
    judge.tracker.addSheet(properties.sheetId, {
      rowCount: properties.gridProperties?.rowCount ?? 1000,
      columnCount: properties.gridProperties?.columnCount ?? 26,
    });
  },
  addTable(add, judge): void {
    if (add.table?.tableId !== undefined) {
      judge.tracker.addTable(add.table.tableId);
    }
  },
  updateSheetProperties(update, judge): void {
    if (masksField(update.fields, "gridProperties")) {
      judge.holdUnmeasured(
        "a tab's grid properties, which can drop rows or columns",
      );
    }
  },
  updateTable(update, judge): void {
    if (masksField(update.fields, "columnProperties")) {
      judge.judgeDropdownReset(update.table ?? {});
    }
  },
  updateCells(update, judge): void {
    if (!masksField(update.fields, "userEnteredValue")) return;
    const rows = update.rows ?? [];
    if (update.range) {
      // A range cell past the rows given is cleared.
      judge.write(
        rangeWrites(update.range, judge, (rowOffset, colOffset) =>
          cellAfter(rows[rowOffset]?.values?.[colOffset]),
        ),
      );
      return;
    }
    if (!update.start) {
      judge.holdUnmeasured("a cell write without a range or start");
      return;
    }
    const start = update.start;
    judge.write(
      rows.flatMap((row, rowOffset) =>
        (row.values ?? []).map((value, colOffset) => ({
          cell: {
            sheetGid: start.sheetId ?? 0,
            rowIndex: (start.rowIndex ?? 0) + rowOffset,
            colIndex: (start.columnIndex ?? 0) + colOffset,
          },
          after: cellAfter(value),
        })),
      ),
    );
  },
  setDataValidation(): void {},
  repeatCell(repeat, judge): void {
    if (!masksField(repeat.fields, "userEnteredValue")) return;
    if (!repeat.range) {
      judge.holdUnmeasured("a repeated cell without a range");
      return;
    }
    const after = cellAfter(repeat.cell);
    judge.write(rangeWrites(repeat.range, judge, () => after));
  },
  pasteData(paste, judge): void {
    if (!valuePasteTypes.has(paste.type ?? "PASTE_NORMAL")) return;
    const coordinate = paste.coordinate;
    if (paste.html || paste.delimiter === undefined || !coordinate) {
      judge.holdUnmeasured("a paste that isn't delimited text at a cell");
      return;
    }
    judge.write(
      pastedFields(paste.data ?? "", paste.delimiter).flatMap(
        (record, rowOffset) =>
          record.map((field, colOffset) => ({
            cell: {
              sheetGid: coordinate.sheetId ?? 0,
              rowIndex: (coordinate.rowIndex ?? 0) + rowOffset,
              colIndex: (coordinate.columnIndex ?? 0) + colOffset,
            },
            after: pastedAfter(field),
          })),
      ),
    );
  },
};

function shiftDimension(dimension: string | undefined): ShiftDimension {
  return dimension === "COLUMNS" ? "COLUMNS" : "ROWS";
}

interface DimensionSpan {
  sheetGid: number;
  dimension: ShiftDimension;
  startIndex: number;
  count: number;
}

function dimensionSpan(
  range: DimensionRange | undefined,
  judge: RequestJudge,
): DimensionSpan | undefined {
  if (range?.endIndex === undefined) {
    judge.holdUnmeasured("rows or columns without an end index");
    return undefined;
  }
  const startIndex = range.startIndex ?? 0;
  return {
    sheetGid: range.sheetId ?? 0,
    dimension: shiftDimension(range.dimension),
    startIndex,
    count: range.endIndex - startIndex,
  };
}

function dimensionGridRange(span: DimensionSpan): GoogleGridRange {
  const endIndex = span.startIndex + span.count;
  if (span.dimension === "ROWS") {
    return {
      sheetId: span.sheetGid,
      startRowIndex: span.startIndex,
      endRowIndex: endIndex,
    };
  }
  return {
    sheetId: span.sheetGid,
    startColumnIndex: span.startIndex,
    endColumnIndex: endIndex,
  };
}

interface RangeShiftRequest {
  range?: GoogleGridRange;
  shiftDimension?: string;
}

interface MeasuredRangeShift {
  bounded: BoundedRange;
  // As an insert counts it; a delete negates the count.
  shift: DimensionShift;
}

function rangeShiftOf(
  request: RangeShiftRequest,
  judge: RequestJudge,
): MeasuredRangeShift | undefined {
  if (!request.range || !request.shiftDimension) {
    judge.holdUnmeasured("a range shift without a range or shift dimension");
    return undefined;
  }
  const bounded = judge.bounded(request.range);
  if (bounded === undefined) return undefined;
  const dimension = shiftDimension(request.shiftDimension);
  const [along, across] =
    dimension === "ROWS"
      ? [bounded.rows, bounded.cols]
      : [bounded.cols, bounded.rows];
  return {
    bounded,
    shift: {
      sheetGid: bounded.sheetGid,
      dimension,
      startIndex: along.startIndex,
      count: along.endIndex - along.startIndex,
      band: across,
    },
  };
}

function rangeWrites(
  range: GoogleGridRange,
  judge: RequestJudge,
  afterAt: (rowOffset: number, colOffset: number) => CellAfter,
): CellWrite[] {
  const bounded = judge.bounded(range);
  if (bounded === undefined) return [];
  return boundedCells(bounded).map((cell) => ({
    cell,
    after: afterAt(
      cell.rowIndex - bounded.rows.startIndex,
      cell.colIndex - bounded.cols.startIndex,
    ),
  }));
}

// A destination a whole number of copies long takes them all; any other takes one.
function tiledSpan(
  span: { startIndex: number; endIndex: number },
  sourceLength: number,
): { startIndex: number; endIndex: number } {
  const length = span.endIndex - span.startIndex;
  const tiledLength =
    length >= sourceLength && length % sourceLength === 0
      ? length
      : sourceLength;
  return {
    startIndex: span.startIndex,
    endIndex: span.startIndex + tiledLength,
  };
}

function isInBounded(range: BoundedRange, cell: GridCell): boolean {
  return (
    cell.sheetGid === range.sheetGid &&
    cell.rowIndex >= range.rows.startIndex &&
    cell.rowIndex < range.rows.endIndex &&
    cell.colIndex >= range.cols.startIndex &&
    cell.colIndex < range.cols.endIndex
  );
}

// An absent mask is held rather than read as touching nothing.
function masksField(fields: string | undefined, field: string): boolean {
  if (fields === undefined) return true;
  return topLevelFields(fields).some((path) => {
    const root = path.split(/[.(]/, 1)[0]?.trim();
    return root === "*" || root === field;
  });
}

function topLevelFields(fields: string): string[] {
  const paths: string[] = [];
  let depth = 0;
  let path = "";
  [...fields].forEach((char) => {
    if (char === "," && depth === 0) {
      paths.push(path);
      path = "";
      return;
    }
    if (char === "(") depth += 1;
    if (char === ")") depth -= 1;
    path += char;
  });
  paths.push(path);
  return paths;
}

function cellAfter(cell: GoogleCellData | undefined): CellAfter {
  const value = cell?.userEnteredValue;
  if (value === undefined || Object.keys(value).length === 0) {
    return { kind: "cleared" };
  }
  return { kind: "value", value };
}

// Google parses pasted text as typed into a cell: a leading "=" is a formula, a numeral a number.
function pastedAfter(field: string): CellAfter {
  if (field === "") return { kind: "cleared" };
  if (field.startsWith("=")) {
    return { kind: "value", value: { formulaValue: field } };
  }
  if (/^-?\d+(\.\d+)?$/.test(field)) {
    return { kind: "value", value: { numberValue: Number(field) } };
  }
  const upper = field.toUpperCase();
  if (upper === "TRUE" || upper === "FALSE") {
    return { kind: "value", value: { boolValue: upper === "TRUE" } };
  }
  return { kind: "value", value: { stringValue: field } };
}

// Records split on newlines and fields on the delimiter, with RFC 4180 quoting.
function pastedFields(data: string, delimiter: string): string[][] {
  const records: string[][] = [];
  let record: string[] = [];
  let field = "";
  let isQuoted = false;
  for (let i = 0; i < data.length; i++) {
    const char = data[i];
    if (isQuoted) {
      if (char !== '"') {
        field += char;
      } else if (data[i + 1] === '"') {
        field += '"';
        i += 1;
      } else {
        isQuoted = false;
      }
      continue;
    }
    if (char === '"' && field === "") {
      isQuoted = true;
      continue;
    }
    if (data.startsWith(delimiter, i)) {
      record.push(field);
      field = "";
      i += delimiter.length - 1;
      continue;
    }
    if (char === "\n") {
      records.push([...record, field]);
      record = [];
      field = "";
      continue;
    }
    field += char;
  }
  records.push([...record, field]);
  return records;
}
