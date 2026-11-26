import type {
  BoundedRange,
  CellBack,
  DimensionShift,
  GridCell,
  IndexSpan,
  TrackerLayer,
} from "../flushClassifierTypes.ts";
import { Grid } from "../Grid.ts";

// An insert or delete of rows or columns, whole or within a band.
export class ShiftLayer implements TrackerLayer {
  private shift: DimensionShift;
  constructor(shift: DimensionShift) {
    this.shift = shift;
  }
  back(cell: GridCell): CellBack {
    const { shift } = this;
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
  forward(range: BoundedRange): BoundedRange | undefined {
    const { shift } = this;
    if (range.sheetGid !== shift.sheetGid) return range;
    const along = Grid.spanAlong(range, shift.dimension);
    const across = Grid.spanAcross(range, shift.dimension);
    if (shift.startIndex >= along.endIndex) return range;
    if (shift.band !== undefined && !Grid.isSpanInside(across, shift.band)) {
      return Grid.doSpansOverlap(across, shift.band) ? undefined : range;
    }
    return Grid.withSpanAlong(range, shift.dimension, this._shiftedSpan(along));
  }
  // An insert at a span's first index moves the whole span; a delete takes the part it covers.
  private _shiftedSpan(span: IndexSpan): IndexSpan {
    const { startIndex, count } = this.shift;
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
