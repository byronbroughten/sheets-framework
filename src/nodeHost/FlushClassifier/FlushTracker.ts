import type {
  BoundedRange,
  GoogleGridRange,
  GridCell,
  GridSize,
  IndexSpan,
  ShiftDimension,
} from "./flushClassifierTypes.ts";
import { Grid } from "./Grid.ts";

export type CellOrigin =
  | { kind: "read"; cell: GridCell }
  // Inserted, appended, on an added tab, or emptied by a cut.
  | { kind: "created" }
  | { kind: "written" }
  // Sorted earlier in the flush, so which read cell it holds is unknown.
  | { kind: "reordered" };

type CellBack = GridCell | Exclude<CellOrigin, { kind: "read" }>;

// One request's effect on the grid, as the cells it left mapped back and the ranges it found mapped on.
interface Layer {
  back(cell: GridCell): CellBack;
  // Undefined once the request split the range, so it no longer stands as one range.
  forward(range: BoundedRange): BoundedRange | undefined;
}

export interface DimensionShift {
  sheetGid: number;
  dimension: ShiftDimension;
  startIndex: number;
  // Positive for an insert, negative for a delete.
  count: number;
  // The span across the other dimension a range shift stays inside; undefined for whole rows or columns.
  band: IndexSpan | undefined;
}

export interface DimensionMove extends IndexSpan {
  sheetGid: number;
  dimension: ShiftDimension;
  // Counted before the moved span is lifted out, as Google counts it.
  destinationIndex: number;
}

// Maps a cell as it stands partway through a flush, or a later flush judged against the same reads, back to the cell the run read, one request at a time.
export class FlushTracker {
  private layers: Layer[];
  private gridSizes: Map<number, GridSize>;
  private addedTableIds: Set<string>;
  // Once a request inserted, deleted, moved or cut cells, a read cell's index may no longer name the cell standing there.
  hasShiftedCells: boolean;
  constructor(gridSizes: Map<number, GridSize>) {
    this.layers = [];
    this.gridSizes = new Map(gridSizes);
    this.addedTableIds = new Set();
    this.hasShiftedCells = false;
  }
  resolve(cell: GridCell): CellOrigin {
    let current = cell;
    for (const layer of [...this.layers].reverse()) {
      const earlier = layer.back(current);
      if ("kind" in earlier) return earlier;
      current = earlier;
    }
    return { kind: "read", cell: current };
  }
  // Where a read range's cells stand now, the inverse of resolve; undefined once a request split them.
  follow(range: BoundedRange): BoundedRange | undefined {
    let current: BoundedRange | undefined = range;
    for (const layer of this.layers) {
      if (current === undefined) return undefined;
      current = layer.forward(current);
    }
    return current;
  }
  // Undefined when an open bound meets a tab whose size no read gave.
  bounded(range: GoogleGridRange): BoundedRange | undefined {
    const sheetGid = range.sheetId ?? 0;
    const size = this.gridSizes.get(sheetGid);
    const endRowIndex = range.endRowIndex ?? size?.rowCount;
    const endColumnIndex = range.endColumnIndex ?? size?.columnCount;
    if (endRowIndex === undefined || endColumnIndex === undefined) {
      return undefined;
    }
    return {
      sheetGid,
      rows: { startIndex: range.startRowIndex ?? 0, endIndex: endRowIndex },
      cols: {
        startIndex: range.startColumnIndex ?? 0,
        endIndex: endColumnIndex,
      },
    };
  }
  isTableAdded(tableId: string): boolean {
    return this.addedTableIds.has(tableId);
  }
  shift(shift: DimensionShift): this {
    if (shift.band === undefined || shift.count > 0) {
      this.grow(shift.sheetGid, shift.dimension, shift.count);
    }
    this.layers.push({
      back: (cell) => shiftedBack(cell, shift),
      forward: (range) => shiftedOn(range, shift),
    });
    this.hasShiftedCells = true;
    return this;
  }
  grow(sheetGid: number, dimension: ShiftDimension, count: number): this {
    const size = this.gridSizes.get(sheetGid);
    if (size !== undefined) {
      this.gridSizes.set(sheetGid, Grid.grown(size, dimension, count));
    }
    return this;
  }
  move(move: DimensionMove): this {
    this.layers.push({
      back: (cell) => movedBack(cell, move),
      forward: (range) => movedOn(range, move),
    });
    this.hasShiftedCells = true;
    return this;
  }
  cut(source: BoundedRange, destination: GridCell): this {
    this.layers.push({
      back: (cell) => cutBack(cell, source, destination),
      forward: (range) => cutOn(range, source, destination),
    });
    this.hasShiftedCells = true;
    return this;
  }
  write(cells: GridCell[]): this {
    const written = new Set(cells.map(cellKey));
    this.layers.push(
      inPlace((cell) =>
        written.has(cellKey(cell)) ? { kind: "written" } : cell,
      ),
    );
    return this;
  }
  reorder(range: GoogleGridRange): this {
    this.layers.push(
      inPlace((cell) =>
        Grid.isInRange(range, cell) ? { kind: "reordered" } : cell,
      ),
    );
    return this;
  }
  addSheet(sheetGid: number, size: GridSize): this {
    this.gridSizes.set(sheetGid, size);
    this.layers.push(
      inPlace((cell) =>
        cell.sheetGid === sheetGid ? { kind: "created" } : cell,
      ),
    );
    return this;
  }
  addTable(tableId: string): this {
    this.addedTableIds.add(tableId);
    return this;
  }
}

function cellKey(cell: GridCell): string {
  return [cell.sheetGid, cell.rowIndex, cell.colIndex].join(":");
}

// A request that changes what cells hold but moves none, so a range stands where it stood.
function inPlace(back: Layer["back"]): Layer {
  return { back, forward: (range) => range };
}

function shiftedBack(cell: GridCell, shift: DimensionShift): CellBack {
  if (cell.sheetGid !== shift.sheetGid) return cell;
  const along = Grid.along(cell, shift.dimension);
  const across = Grid.across(cell, shift.dimension);
  if (shift.band !== undefined && !Grid.isInSpan(across, shift.band)) {
    return cell;
  }
  if (along < shift.startIndex) return cell;
  if (along < shift.startIndex + shift.count) return { kind: "created" };
  return Grid.withAlong(cell, shift.dimension, along - shift.count);
}

// A range shift over only part of the range splits it; one wholly beside or wholly past it leaves it be.
function shiftedOn(
  range: BoundedRange,
  shift: DimensionShift,
): BoundedRange | undefined {
  if (range.sheetGid !== shift.sheetGid) return range;
  const along = Grid.spanAlong(range, shift.dimension);
  const across = Grid.spanAcross(range, shift.dimension);
  if (shift.startIndex >= along.endIndex) return range;
  if (shift.band !== undefined && !Grid.isSpanInside(across, shift.band)) {
    return Grid.doSpansOverlap(across, shift.band) ? undefined : range;
  }
  return Grid.withSpanAlong(range, shift.dimension, shiftedSpan(along, shift));
}

// An insert at a span's first index moves the whole span; a delete takes the part it covers.
function shiftedSpan(span: IndexSpan, shift: DimensionShift): IndexSpan {
  const { startIndex, count } = shift;
  if (count > 0) {
    return {
      startIndex:
        span.startIndex >= startIndex
          ? span.startIndex + count
          : span.startIndex,
      endIndex:
        span.endIndex > startIndex ? span.endIndex + count : span.endIndex,
    };
  }
  const deletedEnd = startIndex - count;
  return {
    startIndex: deletedOn(span.startIndex, startIndex, deletedEnd),
    endIndex: deletedOn(span.endIndex, startIndex, deletedEnd),
  };
}

function deletedOn(
  index: number,
  deletedStart: number,
  deletedEnd: number,
): number {
  if (index < deletedStart) return index;
  if (index < deletedEnd) return deletedStart;
  return index - (deletedEnd - deletedStart);
}

function movedBack(cell: GridCell, move: DimensionMove): GridCell {
  if (cell.sheetGid !== move.sheetGid) return cell;
  const along = Grid.along(cell, move.dimension);
  const movedCount = move.endIndex - move.startIndex;
  const { startIndex, endIndex, destinationIndex } = move;
  if (destinationIndex >= endIndex) {
    const landedStart = destinationIndex - movedCount;
    if (along >= landedStart && along < destinationIndex) {
      return Grid.withAlong(
        cell,
        move.dimension,
        startIndex + along - landedStart,
      );
    }
    if (along >= startIndex && along < landedStart) {
      return Grid.withAlong(cell, move.dimension, along + movedCount);
    }
    return cell;
  }
  if (destinationIndex <= startIndex) {
    if (along >= destinationIndex && along < destinationIndex + movedCount) {
      return Grid.withAlong(
        cell,
        move.dimension,
        startIndex + along - destinationIndex,
      );
    }
    if (along >= destinationIndex + movedCount && along < endIndex) {
      return Grid.withAlong(cell, move.dimension, along - movedCount);
    }
  }
  return cell;
}

// A move carrying part of the range away, or another span into it, splits it.
function movedOn(
  range: BoundedRange,
  move: DimensionMove,
): BoundedRange | undefined {
  if (range.sheetGid !== move.sheetGid) return range;
  const along = Grid.spanAlong(range, move.dimension);
  const indexes = Array.from(
    { length: along.endIndex - along.startIndex },
    (_, offset) => movedOnIndex(along.startIndex + offset, move),
  );
  if (indexes.length === 0) return range;
  const startIndex = Math.min(...indexes);
  const endIndex = Math.max(...indexes) + 1;
  if (endIndex - startIndex !== indexes.length) return undefined;
  return Grid.withSpanAlong(range, move.dimension, { startIndex, endIndex });
}

// The inverse of movedBack along the move's dimension.
function movedOnIndex(index: number, move: DimensionMove): number {
  const { startIndex, endIndex, destinationIndex } = move;
  const movedCount = endIndex - startIndex;
  const isMoved = index >= startIndex && index < endIndex;
  if (destinationIndex >= endIndex) {
    const landedStart = destinationIndex - movedCount;
    if (isMoved) return landedStart + index - startIndex;
    if (index >= endIndex && index < destinationIndex) {
      return index - movedCount;
    }
    return index;
  }
  if (destinationIndex <= startIndex) {
    if (isMoved) return destinationIndex + index - startIndex;
    if (index >= destinationIndex && index < startIndex) {
      return index + movedCount;
    }
  }
  return index;
}

function cutBack(
  cell: GridCell,
  source: BoundedRange,
  destination: GridCell,
): CellBack {
  const rowOffset = cell.rowIndex - destination.rowIndex;
  const colOffset = cell.colIndex - destination.colIndex;
  const isPasted =
    cell.sheetGid === destination.sheetGid &&
    rowOffset >= 0 &&
    rowOffset < source.rows.endIndex - source.rows.startIndex &&
    colOffset >= 0 &&
    colOffset < source.cols.endIndex - source.cols.startIndex;
  if (isPasted) {
    return {
      sheetGid: source.sheetGid,
      rowIndex: source.rows.startIndex + rowOffset,
      colIndex: source.cols.startIndex + colOffset,
    };
  }
  return Grid.isInBounded(source, cell) ? { kind: "created" } : cell;
}

// A cut lifting the whole range carries it to the destination; one lifting only part splits it, and one pasting over it leaves it in place.
function cutOn(
  range: BoundedRange,
  source: BoundedRange,
  destination: GridCell,
): BoundedRange | undefined {
  const isOnSourceTab = range.sheetGid === source.sheetGid;
  if (
    !isOnSourceTab ||
    !Grid.doSpansOverlap(range.rows, source.rows) ||
    !Grid.doSpansOverlap(range.cols, source.cols)
  ) {
    return range;
  }
  if (
    !Grid.isSpanInside(range.rows, source.rows) ||
    !Grid.isSpanInside(range.cols, source.cols)
  ) {
    return undefined;
  }
  const rowOffset = destination.rowIndex - source.rows.startIndex;
  const colOffset = destination.colIndex - source.cols.startIndex;
  return {
    sheetGid: destination.sheetGid,
    rows: offsetSpan(range.rows, rowOffset),
    cols: offsetSpan(range.cols, colOffset),
  };
}

function offsetSpan(span: IndexSpan, offset: number): IndexSpan {
  return {
    startIndex: span.startIndex + offset,
    endIndex: span.endIndex + offset,
  };
}
