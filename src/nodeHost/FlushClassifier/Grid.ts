import type {
  BoundedRange,
  GoogleGridCoordinate,
  GoogleGridRange,
  GridCell,
  GridSize,
  IndexSpan,
  ShiftDimension,
} from "./flushClassifierTypes.ts";

// The fields each dimension reads and writes, so one shift, move or range serves rows and columns.
const dimensionKeys = {
  ROWS: {
    cellIndex: "rowIndex",
    crossCellIndex: "colIndex",
    span: "rows",
    crossSpan: "cols",
    sizeCount: "rowCount",
    rangeStart: "startRowIndex",
    rangeEnd: "endRowIndex",
  },
  COLUMNS: {
    cellIndex: "colIndex",
    crossCellIndex: "rowIndex",
    span: "cols",
    crossSpan: "rows",
    sizeCount: "columnCount",
    rangeStart: "startColumnIndex",
    rangeEnd: "endColumnIndex",
  },
} as const;

export const Grid = {
  cellAt(
    coordinate: GoogleGridCoordinate,
    rowOffset: number,
    colOffset: number,
  ): GridCell {
    return {
      sheetGid: coordinate.sheetId ?? 0,
      rowIndex: (coordinate.rowIndex ?? 0) + rowOffset,
      colIndex: (coordinate.columnIndex ?? 0) + colOffset,
    };
  },
  cells(range: BoundedRange): GridCell[] {
    const { sheetGid, rows, cols } = range;
    return spanIndexes(rows).flatMap((rowIndex) =>
      spanIndexes(cols).map((colIndex) => ({ sheetGid, rowIndex, colIndex })),
    );
  },
  isInSpan(index: number, span: IndexSpan): boolean {
    return index >= span.startIndex && index < span.endIndex;
  },
  isSpanInside(inner: IndexSpan, outer: IndexSpan): boolean {
    return (
      inner.startIndex >= outer.startIndex && inner.endIndex <= outer.endIndex
    );
  },
  doSpansOverlap(first: IndexSpan, second: IndexSpan): boolean {
    return (
      first.startIndex < second.endIndex && second.startIndex < first.endIndex
    );
  },
  isSpanEmpty(span: IndexSpan): boolean {
    return span.endIndex <= span.startIndex;
  },
  doRangesOverlap(first: BoundedRange, second: BoundedRange): boolean {
    return (
      first.sheetGid === second.sheetGid &&
      Grid.doSpansOverlap(first.rows, second.rows) &&
      Grid.doSpansOverlap(first.cols, second.cols)
    );
  },
  isRangeInside(inner: BoundedRange, outer: BoundedRange): boolean {
    return (
      inner.sheetGid === outer.sheetGid &&
      Grid.isSpanInside(inner.rows, outer.rows) &&
      Grid.isSpanInside(inner.cols, outer.cols)
    );
  },
  // Undefined when an open bound meets a tab of unknown size.
  bounded(
    range: GoogleGridRange,
    size: GridSize | undefined,
  ): BoundedRange | undefined {
    const endRowIndex = range.endRowIndex ?? size?.rowCount;
    const endColumnIndex = range.endColumnIndex ?? size?.columnCount;
    if (endRowIndex === undefined || endColumnIndex === undefined) {
      return undefined;
    }
    return {
      sheetGid: range.sheetId ?? 0,
      rows: { startIndex: range.startRowIndex ?? 0, endIndex: endRowIndex },
      cols: {
        startIndex: range.startColumnIndex ?? 0,
        endIndex: endColumnIndex,
      },
    };
  },
  // An absent end is open, as Google reads it.
  isInRange(range: GoogleGridRange, cell: GridCell): boolean {
    return (
      (range.sheetId ?? 0) === cell.sheetGid &&
      Grid.isInSpan(cell.rowIndex, {
        startIndex: range.startRowIndex ?? 0,
        endIndex: range.endRowIndex ?? Infinity,
      }) &&
      Grid.isInSpan(cell.colIndex, {
        startIndex: range.startColumnIndex ?? 0,
        endIndex: range.endColumnIndex ?? Infinity,
      })
    );
  },
  isInBounded(range: BoundedRange, cell: GridCell): boolean {
    return (
      cell.sheetGid === range.sheetGid &&
      Grid.isInSpan(cell.rowIndex, range.rows) &&
      Grid.isInSpan(cell.colIndex, range.cols)
    );
  },
  along(cell: GridCell, dimension: ShiftDimension): number {
    return cell[dimensionKeys[dimension].cellIndex];
  },
  across(cell: GridCell, dimension: ShiftDimension): number {
    return cell[dimensionKeys[dimension].crossCellIndex];
  },
  withAlong(
    cell: GridCell,
    dimension: ShiftDimension,
    index: number,
  ): GridCell {
    return { ...cell, [dimensionKeys[dimension].cellIndex]: index };
  },
  spanAlong(range: BoundedRange, dimension: ShiftDimension): IndexSpan {
    return range[dimensionKeys[dimension].span];
  },
  spanAcross(range: BoundedRange, dimension: ShiftDimension): IndexSpan {
    return range[dimensionKeys[dimension].crossSpan];
  },
  withSpanAlong(
    range: BoundedRange,
    dimension: ShiftDimension,
    span: IndexSpan,
  ): BoundedRange {
    return { ...range, [dimensionKeys[dimension].span]: span };
  },
  grown(size: GridSize, dimension: ShiftDimension, count: number): GridSize {
    const sizeCount = dimensionKeys[dimension].sizeCount;
    return { ...size, [sizeCount]: size[sizeCount] + count };
  },
  // Whole rows or columns: the other dimension stays open.
  dimensionRange(
    sheetGid: number,
    dimension: ShiftDimension,
    span: IndexSpan,
  ): GoogleGridRange {
    const { rangeStart, rangeEnd } = dimensionKeys[dimension];
    return {
      sheetId: sheetGid,
      [rangeStart]: span.startIndex,
      [rangeEnd]: span.endIndex,
    };
  },
};

function spanIndexes(span: IndexSpan): number[] {
  return Array.from(
    { length: Math.max(span.endIndex - span.startIndex, 0) },
    (_, offset) => span.startIndex + offset,
  );
}
