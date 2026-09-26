import type { FakeCell } from "../fakeSheetsService";
import { fakeCells } from "./fakeCells";
import {
  type FakeSheetState,
  fakeSpreadsheet,
  type FakeTableState,
} from "./fakeSpreadsheet";

type GridRange = GoogleAppsScript.Sheets.Schema.GridRange;
export type Dimension = "ROWS" | "COLUMNS";

export interface BoundedRange {
  startRowIndex: number;
  endRowIndex: number;
  startColumnIndex: number;
  endColumnIndex: number;
}

export interface DimensionChange {
  dimension: Dimension;
  startIndex: number;
  count: number;
}

export const fakeGrid = {
  cell(sheet: FakeSheetState, rowIndex: number, colIndex: number): FakeCell {
    return sheet.rows[rowIndex]?.[colIndex] ?? null;
  },
  // The live grid refuses a write past its edge; the fake grows, so thin fixtures stay usable.
  setCell(
    sheet: FakeSheetState,
    rowIndex: number,
    colIndex: number,
    cell: FakeCell,
  ): void {
    while (sheet.rows.length <= rowIndex) sheet.rows.push([]);
    const row = sheet.rows[rowIndex] ?? [];
    while (row.length < colIndex) row.push(null);
    row[colIndex] = cell;
    sheet.rowCount = Math.max(sheet.rowCount, rowIndex + 1);
    sheet.columnCount = Math.max(sheet.columnCount, colIndex + 1);
  },
  updateCell(
    sheet: FakeSheetState,
    rowIndex: number,
    colIndex: number,
    change: (cell: FakeCell) => FakeCell,
  ): void {
    fakeGrid.setCell(
      sheet,
      rowIndex,
      colIndex,
      change(fakeGrid.cell(sheet, rowIndex, colIndex)),
    );
  },
  boundedRange(sheet: FakeSheetState, range: GridRange): BoundedRange {
    return {
      startRowIndex: range.startRowIndex ?? 0,
      endRowIndex: range.endRowIndex ?? sheet.rowCount,
      startColumnIndex: range.startColumnIndex ?? 0,
      endColumnIndex: range.endColumnIndex ?? sheet.columnCount,
    };
  },
  forEachCell(
    range: BoundedRange,
    visit: (rowIndex: number, colIndex: number) => void,
  ): void {
    for (let r = range.startRowIndex; r < range.endRowIndex; r++) {
      for (let c = range.startColumnIndex; c < range.endColumnIndex; c++) {
        visit(r, c);
      }
    }
  },
  // The index just past the last row holding a value, where a sheet append lands.
  endOfData(sheet: FakeSheetState): number {
    const lastIndex = sheet.rows.findLastIndex((row) =>
      row.some((cell) => !fakeCells.isBlank(cell)),
    );
    return lastIndex + 1;
  },
  // Measured live: an insert inside a Table grows it; at its end, only when it inherits from before.
  insert(
    sheet: FakeSheetState,
    change: DimensionChange,
    isInheritingFromBefore: boolean,
  ): void {
    const { dimension, startIndex, count } = change;
    if (dimension === "ROWS") {
      sheet.rows.splice(
        Math.min(startIndex, sheet.rows.length),
        0,
        ...Array.from({ length: count }, (): FakeCell[] => []),
      );
      sheet.rowCount += count;
    } else {
      sheet.rows.forEach((row) => {
        if (row.length <= startIndex) return;
        row.splice(startIndex, 0, ...Array.from({ length: count }, () => null));
      });
      sheet.columnCount += count;
    }
    const shift = insertShift(change, isInheritingFromBefore);
    fakeSpreadsheet.tables(sheet).forEach((table) => {
      const hasGrown = shiftTable(table, dimension, shift);
      if (hasGrown && dimension === "COLUMNS") {
        nameNewColumns(sheet, table, change);
      }
    });
    shiftSheetRanges(sheet, dimension, shift);
  },
  remove(sheet: FakeSheetState, change: DimensionChange): void {
    const { dimension, startIndex, count } = change;
    if (dimension === "ROWS") {
      sheet.rows.splice(startIndex, count);
      sheet.rowCount -= count;
    } else {
      sheet.rows.forEach((row) => row.splice(startIndex, count));
      sheet.columnCount -= count;
    }
    const shift = removeShift(change);
    fakeSpreadsheet.tables(sheet).forEach((table) => {
      shiftTable(table, dimension, shift);
    });
    shiftSheetRanges(sheet, dimension, shift);
  },
};

// Where a dimension change moves a span, and a single index (undefined once removed).
interface DimensionShift {
  span(start: number, end: number): [number, number];
  index(index: number): number | undefined;
}

function insertShift(
  { startIndex, count }: DimensionChange,
  isInheritingFromBefore: boolean,
): DimensionShift {
  return {
    span(start, end) {
      const isInside = isInheritingFromBefore
        ? start < startIndex && startIndex <= end
        : start < startIndex && startIndex < end;
      if (isInside) return [start, end + count];
      if (startIndex <= start) return [start + count, end + count];
      return [start, end];
    },
    index(index) {
      return index >= startIndex ? index + count : index;
    },
  };
}

function removeShift({ startIndex, count }: DimensionChange): DimensionShift {
  function removedBefore(index: number): number {
    return Math.min(Math.max(index - startIndex, 0), count);
  }
  return {
    span(start, end) {
      return [start - removedBefore(start), end - removedBefore(end)];
    },
    index(index) {
      if (index < startIndex) return index;
      if (index < startIndex + count) return undefined;
      return index - count;
    },
  };
}

// Returns whether the Table grew along the changed dimension.
function shiftTable(
  table: FakeTableState,
  dimension: Dimension,
  shift: DimensionShift,
): boolean {
  if (dimension === "ROWS") {
    const [start, end] = shift.span(table.startRowIndex, table.endRowIndex);
    const hasGrown = end - start > table.endRowIndex - table.startRowIndex;
    table.startRowIndex = start;
    table.endRowIndex = end;
    return hasGrown;
  }
  const [start, end] = shift.span(table.startColumnIndex, table.endColumnIndex);
  const hasGrown = end - start > table.endColumnIndex - table.startColumnIndex;
  table.startColumnIndex = start;
  table.endColumnIndex = end;
  table.columnTypes = rekeyColumns(table.columnTypes, shift);
  table.columnValidationValues = rekeyColumns(
    table.columnValidationValues,
    shift,
  );
  table.columnValidationConditionTypes = rekeyColumns(
    table.columnValidationConditionTypes,
    shift,
  );
  return hasGrown;
}

// A column-keyed fact moves with its column and goes when its column does.
function rekeyColumns<FT>(
  byColumn: Record<number, FT> | undefined,
  shift: DimensionShift,
): Record<number, FT> | undefined {
  if (byColumn === undefined) return undefined;
  return Object.entries(byColumn).reduce<Record<number, FT>>(
    (rekeyed, [key, fact]) => {
      const colIndex = shift.index(Number(key));
      if (colIndex !== undefined) rekeyed[colIndex] = fact;
      return rekeyed;
    },
    {},
  );
}

// Measured live: a column a Table grows by is headed "Column <n>", n its 1-based place.
function nameNewColumns(
  sheet: FakeSheetState,
  table: FakeTableState,
  { startIndex, count }: DimensionChange,
): void {
  Array.from({ length: count }, (_, offset) => startIndex + offset).forEach(
    (colIndex) => {
      fakeGrid.setCell(
        sheet,
        table.startRowIndex,
        colIndex,
        `Column ${colIndex - table.startColumnIndex + 1}`,
      );
    },
  );
}

function shiftSheetRanges(
  sheet: FakeSheetState,
  dimension: Dimension,
  shift: DimensionShift,
): void {
  if (dimension === "ROWS") {
    sheet.hiddenRowIndexes = shiftIndexes(sheet.hiddenRowIndexes, shift);
  } else {
    sheet.hiddenColumnIndexes = shiftIndexes(sheet.hiddenColumnIndexes, shift);
  }
  sheet.protectedRanges = sheet.protectedRanges?.flatMap((protection) => {
    const range =
      protection.range === undefined
        ? undefined
        : shiftGridRange(protection.range, dimension, shift);
    return range === undefined ? [] : [{ ...protection, range }];
  });
  sheet.conditionalFormats = sheet.conditionalFormats?.flatMap((rule) => {
    const ranges = (rule.ranges ?? []).flatMap((range) => {
      const shifted = shiftGridRange(range, dimension, shift);
      return shifted === undefined ? [] : [shifted];
    });
    return ranges.length === 0 ? [] : [{ ...rule, ranges }];
  });
}

function shiftIndexes(indexes: number[], shift: DimensionShift): number[] {
  return indexes.flatMap((index) => {
    const shifted = shift.index(index);
    return shifted === undefined ? [] : [shifted];
  });
}

// An unbounded side stays unbounded; a range the change empties is dropped.
function shiftGridRange(
  range: GridRange,
  dimension: Dimension,
  shift: DimensionShift,
): GridRange | undefined {
  const [startKey, endKey] =
    dimension === "ROWS"
      ? (["startRowIndex", "endRowIndex"] as const)
      : (["startColumnIndex", "endColumnIndex"] as const);
  const end = range[endKey];
  const [start, shiftedEnd] = shift.span(
    range[startKey] ?? 0,
    end ?? Number.MAX_SAFE_INTEGER,
  );
  if (end !== undefined && shiftedEnd <= start) return undefined;
  return {
    ...range,
    ...(range[startKey] === undefined ? {} : { [startKey]: start }),
    ...(end === undefined ? {} : { [endKey]: shiftedEnd }),
  };
}
