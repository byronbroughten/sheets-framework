import type {
  BoundedRange,
  CellOrigin,
  DimensionMove,
  DimensionShift,
  GoogleGridRange,
  GridCell,
  GridSize,
  ShiftDimension,
  TrackerLayer,
} from "./flushClassifierTypes.ts";
import { CutLayer } from "./FlushTracker/CutLayer.ts";
import { MoveLayer } from "./FlushTracker/MoveLayer.ts";
import { ShiftLayer } from "./FlushTracker/ShiftLayer.ts";
import { Grid } from "./Grid.ts";

// Maps a cell as it stands partway through a flush, or a later flush judged against the same reads, back to the cell the run read, one request at a time.
export class FlushTracker {
  private layers: TrackerLayer[];
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
    return Grid.bounded(range, this.gridSizes.get(range.sheetId ?? 0));
  }
  isTableAdded(tableId: string): boolean {
    return this.addedTableIds.has(tableId);
  }
  shift(shift: DimensionShift): this {
    if (shift.band === undefined || shift.count > 0) {
      this.grow(shift.sheetGid, shift.dimension, shift.count);
    }
    this.layers.push(new ShiftLayer(shift));
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
    this.layers.push(new MoveLayer(move));
    this.hasShiftedCells = true;
    return this;
  }
  cut(source: BoundedRange, destination: GridCell): this {
    this.layers.push(new CutLayer({ source, destination }));
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
function inPlace(back: TrackerLayer["back"]): TrackerLayer {
  return { back, forward: (range) => range };
}
