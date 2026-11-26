import type {
  BoundedRange,
  CellBack,
  GridCell,
  IndexSpan,
  RangeCut,
  TrackerLayer,
} from "../flushClassifierTypes.ts";
import { Grid } from "../Grid.ts";

// A cut of a range pasted at a destination cell.
export class CutLayer implements TrackerLayer {
  private cut: RangeCut;
  constructor(cut: RangeCut) {
    this.cut = cut;
  }
  back(cell: GridCell): CellBack {
    const { source, destination } = this.cut;
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
  forward(range: BoundedRange): BoundedRange | undefined {
    const { source, destination } = this.cut;
    if (!Grid.doRangesOverlap(range, source)) return range;
    if (!Grid.isRangeInside(range, source)) return undefined;
    const rowOffset = destination.rowIndex - source.rows.startIndex;
    const colOffset = destination.colIndex - source.cols.startIndex;
    return {
      sheetGid: destination.sheetGid,
      rows: offsetSpan(range.rows, rowOffset),
      cols: offsetSpan(range.cols, colOffset),
    };
  }
}

function offsetSpan(span: IndexSpan, offset: number): IndexSpan {
  return {
    startIndex: span.startIndex + offset,
    endIndex: span.endIndex + offset,
  };
}
