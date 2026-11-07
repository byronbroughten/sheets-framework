import type {
  GoogleGridRange,
  GridCell,
  GridSize,
} from "./flushClassifierTypes";
import { isInRange } from "./RunReads";

export type CellOrigin =
  | { kind: "read"; cell: GridCell }
  // Inserted, appended, on an added tab, or emptied by a cut.
  | { kind: "created" }
  | { kind: "written" }
  // Sorted earlier in the flush, so which read cell it holds is unknown.
  | { kind: "reordered" };

type Layer = (
  cell: GridCell,
) => GridCell | Exclude<CellOrigin, { kind: "read" }>;

export type ShiftDimension = "ROWS" | "COLUMNS";

interface IndexSpan {
  startIndex: number;
  endIndex: number;
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

export interface BoundedRange {
  sheetGid: number;
  rows: IndexSpan;
  cols: IndexSpan;
}

// Maps a cell as it stands partway through a flush back to the cell the run read, one request at a time.
export class FlushTracker {
  private layers: Layer[];
  private gridSizes: Map<number, GridSize>;
  private addedTableIds: Set<string>;
  constructor(gridSizes: Map<number, GridSize>) {
    this.layers = [];
    this.gridSizes = new Map(gridSizes);
    this.addedTableIds = new Set();
  }
  resolve(cell: GridCell): CellOrigin {
    let current = cell;
    for (const layer of [...this.layers].reverse()) {
      const earlier = layer(current);
      if ("kind" in earlier) return earlier;
      current = earlier;
    }
    return { kind: "read", cell: current };
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
      this._grow(shift.sheetGid, shift.dimension, shift.count);
    }
    this.layers.push((cell) => shiftedBack(cell, shift));
    return this;
  }
  grow(sheetGid: number, dimension: ShiftDimension, count: number): this {
    this._grow(sheetGid, dimension, count);
    return this;
  }
  move(move: DimensionMove): this {
    this.layers.push((cell) => movedBack(cell, move));
    return this;
  }
  cut(source: BoundedRange, destination: GridCell): this {
    this.layers.push((cell) => cutBack(cell, source, destination));
    return this;
  }
  write(cells: GridCell[]): this {
    const written = new Set(cells.map(cellKey));
    this.layers.push((cell) =>
      written.has(cellKey(cell)) ? { kind: "written" } : cell,
    );
    return this;
  }
  reorder(range: GoogleGridRange): this {
    this.layers.push((cell) =>
      isInRange(range, cell) ? { kind: "reordered" } : cell,
    );
    return this;
  }
  addSheet(sheetGid: number, size: GridSize): this {
    this.gridSizes.set(sheetGid, size);
    this.layers.push((cell) =>
      cell.sheetGid === sheetGid ? { kind: "created" } : cell,
    );
    return this;
  }
  addTable(tableId: string): this {
    this.addedTableIds.add(tableId);
    return this;
  }
  private _grow(
    sheetGid: number,
    dimension: ShiftDimension,
    count: number,
  ): void {
    const size = this.gridSizes.get(sheetGid);
    if (size === undefined) return;
    if (dimension === "ROWS") {
      this.gridSizes.set(sheetGid, {
        ...size,
        rowCount: size.rowCount + count,
      });
    } else {
      this.gridSizes.set(sheetGid, {
        ...size,
        columnCount: size.columnCount + count,
      });
    }
  }
}

function cellKey(cell: GridCell): string {
  return [cell.sheetGid, cell.rowIndex, cell.colIndex].join(":");
}

function isInSpan(index: number, span: IndexSpan): boolean {
  return index >= span.startIndex && index < span.endIndex;
}

// The index along a dimension and the one across it, so one shift or move serves rows and columns.
function alongAndAcross(
  cell: GridCell,
  dimension: ShiftDimension,
): [number, number] {
  if (dimension === "ROWS") return [cell.rowIndex, cell.colIndex];
  return [cell.colIndex, cell.rowIndex];
}

function withAlong(
  cell: GridCell,
  dimension: ShiftDimension,
  index: number,
): GridCell {
  if (dimension === "ROWS") return { ...cell, rowIndex: index };
  return { ...cell, colIndex: index };
}

function shiftedBack(cell: GridCell, shift: DimensionShift): ReturnType<Layer> {
  if (cell.sheetGid !== shift.sheetGid) return cell;
  const [along, across] = alongAndAcross(cell, shift.dimension);
  if (shift.band !== undefined && !isInSpan(across, shift.band)) return cell;
  if (along < shift.startIndex) return cell;
  if (along < shift.startIndex + shift.count) return { kind: "created" };
  return withAlong(cell, shift.dimension, along - shift.count);
}

function movedBack(cell: GridCell, move: DimensionMove): GridCell {
  if (cell.sheetGid !== move.sheetGid) return cell;
  const [along] = alongAndAcross(cell, move.dimension);
  const movedCount = move.endIndex - move.startIndex;
  const { startIndex, endIndex, destinationIndex } = move;
  if (destinationIndex >= endIndex) {
    const landedStart = destinationIndex - movedCount;
    if (along >= landedStart && along < destinationIndex) {
      return withAlong(cell, move.dimension, startIndex + along - landedStart);
    }
    if (along >= startIndex && along < landedStart) {
      return withAlong(cell, move.dimension, along + movedCount);
    }
    return cell;
  }
  if (destinationIndex <= startIndex) {
    if (along >= destinationIndex && along < destinationIndex + movedCount) {
      return withAlong(
        cell,
        move.dimension,
        startIndex + along - destinationIndex,
      );
    }
    if (along >= destinationIndex + movedCount && along < endIndex) {
      return withAlong(cell, move.dimension, along - movedCount);
    }
  }
  return cell;
}

function cutBack(
  cell: GridCell,
  source: BoundedRange,
  destination: GridCell,
): ReturnType<Layer> {
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
  const isVacated =
    cell.sheetGid === source.sheetGid &&
    isInSpan(cell.rowIndex, source.rows) &&
    isInSpan(cell.colIndex, source.cols);
  return isVacated ? { kind: "created" } : cell;
}
