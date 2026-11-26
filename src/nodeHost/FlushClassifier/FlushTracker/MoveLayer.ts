import type {
  BoundedRange,
  DimensionMove,
  GridCell,
  TrackerLayer,
} from "../flushClassifierTypes.ts";
import { Grid } from "../Grid.ts";

// A move of whole rows or columns to another index.
export class MoveLayer implements TrackerLayer {
  private move: DimensionMove;
  constructor(move: DimensionMove) {
    this.move = move;
  }
  back(cell: GridCell): GridCell {
    const { move } = this;
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
  forward(range: BoundedRange): BoundedRange | undefined {
    const { move } = this;
    if (range.sheetGid !== move.sheetGid) return range;
    const along = Grid.spanAlong(range, move.dimension);
    const indexes = Array.from(
      { length: along.endIndex - along.startIndex },
      (_, offset) => this._movedOnIndex(along.startIndex + offset),
    );
    if (indexes.length === 0) return range;
    const startIndex = Math.min(...indexes);
    const endIndex = Math.max(...indexes) + 1;
    if (endIndex - startIndex !== indexes.length) return undefined;
    return Grid.withSpanAlong(range, move.dimension, { startIndex, endIndex });
  }
  // The inverse of back along the move's dimension.
  private _movedOnIndex(index: number): number {
    const { startIndex, endIndex, destinationIndex } = this.move;
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
}
