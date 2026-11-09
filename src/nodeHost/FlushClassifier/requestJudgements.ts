import type {
  GoogleRequest,
  ModeledRequestVerb,
} from "../../00_Source/GoogleSheets/GoogleSheetsAPI";
import { Rfc4180 } from "../Rfc4180";
import type {
  BoundedRange,
  CellAfter,
  GoogleGridRange,
  IndexSpan,
  ShiftDimension,
} from "./flushClassifierTypes";
import type { DimensionShift, FlushTracker } from "./FlushTracker";
import { Grid } from "./Grid";
import type { CellWrite, RequestJudge } from "./RequestJudge";
import { cellAfter } from "./requestJudgements/cellAfter";
import { fieldMask } from "./requestJudgements/fieldMask";

type DimensionRange = GoogleAppsScript.Sheets.Schema.DimensionRange;

// cutPaste and moveDimension aren't built yet, but #219's column move sends cutPaste, so both are judged ahead of their builders.
export type ClassifiedVerb = ModeledRequestVerb | "cutPaste" | "moveDimension";

type RequestJudgement<CV extends ClassifiedVerb> = (
  inner: NonNullable<GoogleRequest[CV]>,
  judge: RequestJudge,
  tracker: FlushTracker,
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
  appendDimension(append, _judge, tracker): void {
    tracker.grow(
      append.sheetId ?? 0,
      shiftDimension(append.dimension),
      append.length ?? 0,
    );
  },
  insertDimension(insert, judge, tracker): void {
    const range = dimensionSpan(insert.range, judge);
    if (range === undefined) return;
    tracker.shift({ ...range, band: undefined });
  },
  deleteDimension(remove, judge, tracker): void {
    const range = dimensionSpan(remove.range, judge);
    if (range === undefined) return;
    judge.delete(
      judge.cellsIn(
        Grid.dimensionRange(range.sheetGid, range.dimension, {
          startIndex: range.startIndex,
          endIndex: range.startIndex + range.count,
        }),
      ),
    );
    tracker.shift({ ...range, count: -range.count, band: undefined });
  },
  insertRange(insert, judge, tracker): void {
    const measured = rangeShiftOf(insert, judge);
    if (measured === undefined) return;
    tracker.shift(measured.shift);
  },
  deleteRange(remove, judge, tracker): void {
    const measured = rangeShiftOf(remove, judge);
    if (measured === undefined) return;
    judge.delete(Grid.cells(measured.bounded));
    tracker.shift({ ...measured.shift, count: -measured.shift.count });
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
    const afters = Grid.cells(pasted).map((cell) => ({
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
  cutPaste(cut, judge, tracker): void {
    const coordinate = cut.destination;
    if (!cut.source || !coordinate) {
      judge.holdUnmeasured("a cut without a source or destination");
      return;
    }
    const source = judge.bounded(cut.source);
    if (!source) return;
    // Google cuts all the source whatever the paste type, so a paste without values empties it and leaves the destination's values.
    if (!valuePasteTypes.has(cut.pasteType ?? "PASTE_NORMAL")) {
      judge.write(
        Grid.cells(source).map((cell) => ({ cell, after: clearedAfter })),
      );
      return;
    }
    // A cell the cut lands on that it also lifts from loses nothing: its content moves with the rest.
    const overwritten = Grid.cells(source)
      .map((from) => ({
        from,
        cell: Grid.cellAt(
          coordinate,
          from.rowIndex - source.rows.startIndex,
          from.colIndex - source.cols.startIndex,
        ),
      }))
      .filter(({ cell }) => !Grid.isInBounded(source, cell));
    judge.write(
      overwritten.map(({ from, cell }) => ({
        cell,
        after: judge.afterFrom(from),
      })),
    );
    tracker.cut(source, Grid.cellAt(coordinate, 0, 0));
  },
  moveDimension(move, judge, tracker): void {
    const range = dimensionSpan(move.source, judge);
    if (range === undefined) return;
    if (move.destinationIndex === undefined) {
      judge.holdUnmeasured("a move without a destination");
      return;
    }
    tracker.move({
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
  sortRange(sort, _judge, tracker): void {
    if (sort.range) tracker.reorder(sort.range);
  },
  addConditionalFormatRule(): void {},
  deleteConditionalFormatRule(): void {},
  addProtectedRange(): void {},
  deleteProtectedRange(): void {},
  addSheet(add, _judge, tracker): void {
    const properties = add.properties;
    // A tab Google numbers itself can't be named by a later request in the same flush.
    if (properties?.sheetId === undefined) return;
    tracker.addSheet(properties.sheetId, {
      rowCount: properties.gridProperties?.rowCount ?? 1000,
      columnCount: properties.gridProperties?.columnCount ?? 26,
    });
  },
  addTable(add, _judge, tracker): void {
    if (add.table?.tableId !== undefined) {
      tracker.addTable(add.table.tableId);
    }
  },
  updateSheetProperties(update, judge): void {
    if (fieldMask.masks(update.fields, "gridProperties")) {
      judge.holdUnmeasured(
        "a tab's grid properties, which can drop rows or columns",
      );
    }
  },
  updateTable(update, judge): void {
    if (fieldMask.masks(update.fields, "columnProperties")) {
      judge.judgeDropdownReset(update.table ?? {});
    }
  },
  updateCells(update, judge): void {
    if (!fieldMask.masks(update.fields, "userEnteredValue")) return;
    const rows = update.rows ?? [];
    if (update.range) {
      // A range cell past the rows given is cleared.
      judge.write(
        rangeWrites(update.range, judge, (rowOffset, colOffset) =>
          cellAfter.written(rows[rowOffset]?.values?.[colOffset]),
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
          cell: Grid.cellAt(start, rowOffset, colOffset),
          after: cellAfter.written(value),
        })),
      ),
    );
  },
  setDataValidation(): void {},
  repeatCell(repeat, judge): void {
    if (!fieldMask.masks(repeat.fields, "userEnteredValue")) return;
    if (!repeat.range) {
      judge.holdUnmeasured("a repeated cell without a range");
      return;
    }
    const after = cellAfter.written(repeat.cell);
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
      Rfc4180.records(paste.data ?? "", paste.delimiter).flatMap(
        (record, rowOffset) =>
          record.map((field, colOffset) => ({
            cell: Grid.cellAt(coordinate, rowOffset, colOffset),
            after: cellAfter.pasted(field),
          })),
      ),
    );
  },
};

const clearedAfter: CellAfter = { kind: "cleared" };

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
  const along = Grid.spanAlong(bounded, dimension);
  return {
    bounded,
    shift: {
      sheetGid: bounded.sheetGid,
      dimension,
      startIndex: along.startIndex,
      count: along.endIndex - along.startIndex,
      band: Grid.spanAcross(bounded, dimension),
    },
  };
}

function rangeWrites(
  range: GoogleGridRange,
  judge: RequestJudge,
  afterAt: (rowOffset: number, colOffset: number) => CellAfter,
): CellWrite[] {
  return judge.cellsIn(range).map((cell) => ({
    cell,
    after: afterAt(
      cell.rowIndex - (range.startRowIndex ?? 0),
      cell.colIndex - (range.startColumnIndex ?? 0),
    ),
  }));
}

// A destination a whole number of copies long takes them all; any other takes one.
function tiledSpan(span: IndexSpan, sourceLength: number): IndexSpan {
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
