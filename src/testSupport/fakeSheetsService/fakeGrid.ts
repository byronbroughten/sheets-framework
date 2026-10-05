import type { FakeCell } from "../fakeSheetsService";
import { fakeCells } from "./fakeCells";
import { FakeGoogleRefusal } from "./FakeGoogleRefusal";
import { type FakeSheetState, type FakeTableState } from "./fakeSpreadsheet";

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

// A change bounded across its dimension, e.g. to some columns for a ROWS shift.
export interface BandChange extends DimensionChange {
  crossStartIndex: number;
  crossEndIndex: number;
}

const partOfTableRefusal =
  "You cannot insert or delete cells over part of a table.";

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
  // Measured live: a range starting past the grid's edge is refused, and one running past it is clipped.
  rangeInGrid(sheet: FakeSheetState, range: GridRange): BoundedRange {
    const bounded = fakeGrid.boundedRange(sheet, range);
    if (
      bounded.startRowIndex >= sheet.rowCount ||
      bounded.startColumnIndex >= sheet.columnCount
    ) {
      throw new FakeGoogleRefusal(
        `Range ('${sheet.title}'!${a1Range(bounded)}) exceeds grid limits. Max rows: ${sheet.rowCount}, max columns: ${sheet.columnCount}`,
      );
    }
    return {
      ...bounded,
      endRowIndex: Math.min(bounded.endRowIndex, sheet.rowCount),
      endColumnIndex: Math.min(bounded.endColumnIndex, sheet.columnCount),
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
        let neighbour: FakeCell | undefined;
        if (isInheritingFromBefore) neighbour = row[startIndex - 1];
        if (row.length <= startIndex && neighbour === undefined) return;
        row.splice(
          startIndex,
          0,
          ...Array.from({ length: count }, () =>
            fakeCells.inherited(neighbour),
          ),
        );
      });
      sheet.columnCount += count;
    }
    const shift = insertShift(change, isInheritingFromBefore);
    sheet.tables.forEach((table) => {
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
    sheet.tables.forEach((table) => {
      shiftTable(table, dimension, shift);
    });
    shiftSheetRanges(sheet, dimension, shift);
  },
  // Measured live: only the band's cells move, and a Table it would split is refused (sheets-framework#53, #56, #59).
  insertBand(sheet: FakeSheetState, change: BandChange): void {
    const shift = insertShift(change, false);
    validateBandTables(sheet, change);
    moveBandCells(sheet, change, (line) => {
      line.splice(
        Math.min(change.startIndex, line.length),
        0,
        ...Array.from({ length: change.count }, (): FakeCell => null),
      );
    });
    bandTables(sheet, change).forEach((table) => {
      const hasGrown = shiftTable(table, change.dimension, shift);
      if (hasGrown && change.dimension === "COLUMNS") {
        nameNewColumns(sheet, table, change);
      }
    });
    shiftBandRanges(sheet, change, shift);
    growGridToContent(sheet, change.dimension);
  },
  // Measured live: the grid keeps its size, and a Table may shrink to its header but never lose it.
  removeBand(sheet: FakeSheetState, change: BandChange): void {
    validateBandTables(sheet, change);
    validateNoHeaderRemoved(sheet, change);
    moveBandCells(sheet, change, (line) => {
      line.splice(change.startIndex, change.count);
    });
    const shift = removeShift(change);
    bandTables(sheet, change).forEach((table) => {
      shiftTable(table, change.dimension, shift);
    });
    shiftBandRanges(sheet, change, shift);
  },
};

function a1Range(range: BoundedRange): string {
  return `${columnLetters(range.startColumnIndex)}${range.startRowIndex + 1}:${columnLetters(range.endColumnIndex - 1)}${range.endRowIndex}`;
}

function columnLetters(colIndex: number): string {
  const letter = String.fromCharCode(65 + (colIndex % 26));
  if (colIndex < 26) return letter;
  return columnLetters(Math.floor(colIndex / 26) - 1) + letter;
}

function validateBandTables(sheet: FakeSheetState, change: BandChange): void {
  sheet.tables.forEach((table) => {
    const [, end] = tableSpan(table, change.dimension);
    const [crossStart, crossEnd] = tableSpan(table, crossDimension(change));
    const isReached = end > change.startIndex;
    if (isReached && bandOverlap(change, crossStart, crossEnd) === "part") {
      throw new FakeGoogleRefusal(partOfTableRefusal);
    }
  });
}

function validateNoHeaderRemoved(
  sheet: FakeSheetState,
  change: BandChange,
): void {
  if (change.dimension !== "ROWS") return;
  const removesHeader = bandTables(sheet, change).some(
    (table) =>
      change.startIndex <= table.startRowIndex &&
      table.startRowIndex < change.startIndex + change.count,
  );
  if (removesHeader) {
    throw new FakeGoogleRefusal(
      "Cannot delete a table header row. Consider hiding the row instead.",
    );
  }
}

// The Tables a band holds whole across its dimension; validateBandTables refuses the rest it reaches.
function bandTables(
  sheet: FakeSheetState,
  change: BandChange,
): FakeTableState[] {
  return sheet.tables.filter((table) => {
    const [crossStart, crossEnd] = tableSpan(table, crossDimension(change));
    return bandOverlap(change, crossStart, crossEnd) === "whole";
  });
}

function tableSpan(
  table: FakeTableState,
  dimension: Dimension,
): [number, number] {
  if (dimension === "ROWS") return [table.startRowIndex, table.endRowIndex];
  return [table.startColumnIndex, table.endColumnIndex];
}

function crossDimension(change: BandChange): Dimension {
  return change.dimension === "ROWS" ? "COLUMNS" : "ROWS";
}

function bandOverlap(
  change: BandChange,
  crossStart: number,
  crossEnd: number,
): "none" | "part" | "whole" {
  if (crossEnd <= change.crossStartIndex) return "none";
  if (crossStart >= change.crossEndIndex) return "none";
  if (
    change.crossStartIndex <= crossStart &&
    crossEnd <= change.crossEndIndex
  ) {
    return "whole";
  }
  return "part";
}

// Each line runs along the shift, one per index across the band.
function moveBandCells(
  sheet: FakeSheetState,
  change: BandChange,
  edit: (line: FakeCell[]) => void,
): void {
  indexes(change.crossStartIndex, change.crossEndIndex).forEach((cross) => {
    const line = bandLine(sheet, change.dimension, cross);
    const lineLength = line.length;
    edit(line);
    indexes(0, Math.max(lineLength, line.length)).forEach((along) => {
      const [rowIndex, colIndex] =
        change.dimension === "ROWS" ? [along, cross] : [cross, along];
      placeCell(sheet, rowIndex, colIndex, line[along] ?? null);
    });
  });
}

function indexes(start: number, end: number): number[] {
  return Array.from({ length: Math.max(0, end - start) }, (_, i) => start + i);
}

function bandLine(
  sheet: FakeSheetState,
  dimension: Dimension,
  cross: number,
): FakeCell[] {
  if (dimension === "ROWS") {
    return sheet.rows.map((row) => row[cross] ?? null);
  }
  return [...(sheet.rows[cross] ?? [])];
}

// Unlike setCell, it leaves the grid's size alone; growGridToContent settles that after the move.
function placeCell(
  sheet: FakeSheetState,
  rowIndex: number,
  colIndex: number,
  cell: FakeCell,
): void {
  const row = sheet.rows[rowIndex];
  if (cell === null && (row === undefined || row.length <= colIndex)) return;
  while (sheet.rows.length <= rowIndex) sheet.rows.push([]);
  const target = sheet.rows[rowIndex] ?? [];
  while (target.length < colIndex) target.push(null);
  target[colIndex] = cell;
}

// Measured live: an insert grows the grid only as far as the cells it pushes past the edge.
function growGridToContent(sheet: FakeSheetState, dimension: Dimension): void {
  if (dimension === "ROWS") {
    const lastFilled = sheet.rows.findLastIndex((row) =>
      row.some((cell) => cell !== null),
    );
    sheet.rowCount = Math.max(
      sheet.rowCount,
      lastFilled + 1,
      ...sheet.tables.map((table) => table.endRowIndex),
    );
    return;
  }
  sheet.columnCount = Math.max(
    sheet.columnCount,
    ...sheet.rows.map((row) => row.findLastIndex((cell) => cell !== null) + 1),
    ...sheet.tables.map((table) => table.endColumnIndex),
  );
}

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

// Rules and protections the band holds whole move with it; one it cuts across is beyond what was measured.
function shiftBandRanges(
  sheet: FakeSheetState,
  change: BandChange,
  shift: DimensionShift,
): void {
  function shifted(range: GridRange, owner: string): GridRange | undefined {
    if (isWholeSheetRange(range)) return range;
    const [crossStart, crossEnd] = gridRangeSpan(range, crossDimension(change));
    const [, end] = gridRangeSpan(range, change.dimension);
    const overlap = bandOverlap(change, crossStart, crossEnd);
    if (overlap === "none" || end <= change.startIndex) return range;
    if (overlap === "part") {
      throw new Error(
        `The fake Sheets service does not replay a range shift over part of a ${owner}.`,
      );
    }
    return shiftGridRange(range, change.dimension, shift);
  }
  sheet.protectedRanges = sheet.protectedRanges?.flatMap((protection) => {
    const range =
      protection.range === undefined
        ? undefined
        : shifted(protection.range, "protected range");
    return range === undefined ? [] : [{ ...protection, range }];
  });
  sheet.conditionalFormats = sheet.conditionalFormats?.flatMap((rule) => {
    const ranges = (rule.ranges ?? []).flatMap((range) => {
      const moved = shifted(range, "conditional-format rule");
      return moved === undefined ? [] : [moved];
    });
    return ranges.length === 0 ? [] : [{ ...rule, ranges }];
  });
}

// A whole-sheet range has no coordinates for a shift to move or cut.
function isWholeSheetRange(range: GridRange): boolean {
  return (
    range.startRowIndex === undefined &&
    range.endRowIndex === undefined &&
    range.startColumnIndex === undefined &&
    range.endColumnIndex === undefined
  );
}

function gridRangeSpan(
  range: GridRange,
  dimension: Dimension,
): [number, number] {
  if (dimension === "ROWS") {
    return [
      range.startRowIndex ?? 0,
      range.endRowIndex ?? Number.MAX_SAFE_INTEGER,
    ];
  }
  return [
    range.startColumnIndex ?? 0,
    range.endColumnIndex ?? Number.MAX_SAFE_INTEGER,
  ];
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
