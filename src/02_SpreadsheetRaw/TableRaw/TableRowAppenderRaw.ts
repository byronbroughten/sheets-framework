import type {
  BoundedGridRange,
  CopyPasteOperation,
  CopyPasteType,
  InsertRangeOperation,
} from "../../00_Source/RawSource/RawSource.ts";
import { SheetIndex } from "../../00_Source/RawSource/SheetIndex.ts";
import { columnRun, TableBaseRaw } from "../ClassBases/TableBaseRaw.ts";
import type { RowStateRaw } from "../ClassTypes/StateRaw.ts";
import { TableRaw } from "../TableRaw.ts";
import { TableColumnTypeGathererRaw } from "./TableColumnTypeGathererRaw.ts";

export class TableRowAppenderRaw extends TableBaseRaw {
  get table(): TableRaw {
    return new TableRaw(this.tableRawProps);
  }
  private get columnTypeGatherer(): TableColumnTypeGathererRaw {
    return new TableColumnTypeGathererRaw(this.tableRawProps);
  }
  // Before the Table-end column inserts are counted, since growth is sent ahead of them.
  gatherAppendTableRowsOperation(): void {
    const { table } = this;
    const appendedRowIndexes = this._queuedRowAppendIndexes();
    if (appendedRowIndexes.length === 0) return;
    const appendedRowCount = appendedRowIndexes.length;
    // Not dataRowCount: a same-run re-fetch resets the Table's end but keeps the queued appends.
    const modelRow = table.dataRowGridRange(
      Math.min(...appendedRowIndexes) - 1,
    );
    const newRows: BoundedGridRange = {
      ...modelRow,
      startRowIndex: modelRow.endRowIndex,
      endRowIndex: SheetIndex.row(modelRow.endRowIndex + appendedRowCount),
    };
    // The appendDimension sent first already grows a Table ending on the grid's last row; an insert would push that row out of it.
    // Less the rows the grid grew it by, which that appendDimension also sends.
    const endsOnGridEdge =
      newRows.startRowIndex - table.writes.gridGrownRowCount ===
      table.sheet.rowCountBeforeFlush;
    table.sheet.queueGridRowsThrough(newRows.endRowIndex);
    const insert: InsertRangeOperation[] = endsOnGridEdge
      ? []
      : [{ kind: "insertRange", range: newRows, shiftDimension: "ROWS" }];
    table.writeOperations.appendTableRows.push({
      tableId: table.tableId,
      newRows,
      operations: [
        ...insert,
        {
          kind: "updateTableRange",
          tableId: table.tableId,
          range: { ...newRows, startRowIndex: table.startRowIndex },
        },
        {
          kind: "copyPaste",
          source: modelRow,
          destination: newRows,
          pasteType: "PASTE_FORMAT",
        },
        ...columnRunCopies(
          modelRow,
          newRows,
          this._untypedColumnRuns(),
          "PASTE_DATA_VALIDATION",
        ),
        ...columnRunCopies(
          modelRow,
          newRows,
          this._appendedFormulaColumnRuns(),
          "PASTE_FORMULA",
        ),
      ],
    });
  }
  // An appended row deleted this flush still grows, then goes with the deletes, so the rows after it keep their place.
  private _queuedRowAppendIndexes(): number[] {
    return Array.from(this.table.rowWrites).flatMap(([rowIndex, writes]) =>
      writes.appendRow ? [rowIndex] : [],
    );
  }
  // A typed column's validation is the Table's own, so only Automatic columns get theirs copied.
  private _untypedColumnRuns(): ColumnRun[] {
    const columnTypes = this.columnTypeGatherer.columnTypesAfterFlush();
    return columnRuns(
      this.table.fullTableColIndexes.filter(
        (colIndex) => columnTypes.get(colIndex) === undefined,
      ),
    );
  }
  // Only the named columns, so the paste never reaches a plain value the append left unwritten.
  private _appendedFormulaColumnRuns(): ColumnRun[] {
    const { table } = this;
    const { appendedFormulaColIndexes } = table.writes;
    return columnRuns(
      table.fullTableColIndexes.filter((colIndex) =>
        appendedFormulaColIndexes.has(colIndex),
      ),
    );
  }
  // Matches Sheets, which grows a Table ending on the grid's last row with the grid; one with appends queued sets its own range as they gather.
  growWithGridEdge(addedRowCount: number): void {
    if (!this._endsOnWorkingGridEdge) return;
    if (this._queuedRowAppendIndexes().length > 0) return;
    const { table } = this;
    const lastRowIndex = table.dataRowCount - 1;
    const lastRowState =
      lastRowIndex < 0 ? undefined : this.rowStates.get(lastRowIndex);
    for (let offset = 1; offset <= addedRowCount; offset++) {
      table.growDataRowCount();
      if (lastRowState === undefined) continue;
      this.rowStates.set(
        lastRowIndex + offset,
        this._carriedDownRowState(lastRowState),
      );
    }
    table.writes.gridGrownRowCount += addedRowCount;
  }
  // A Table whose rows have moved can't say where it ends, and every row read on it throws anyway.
  private get _endsOnWorkingGridEdge(): boolean {
    const properties = this.tableProperties;
    if (properties === undefined || properties.rowIndexesAreStale) return false;
    return properties.endRowIndex === this.table.sheet.rowCount;
  }
  // As Sheets fills a grown row: each formula column's formula carried down, every other cell blank.
  // Sheets and the fake decide per cell; row state holds no per-cell formula fact, so a column's sampled top row stands in for it.
  private _carriedDownRowState(lastRowState: RowStateRaw): RowStateRaw {
    return new Map(
      Array.from(lastRowState, ([colIndex, { value }]) => {
        const { isSampledFormula } = this.table.column(colIndex).profile;
        return [colIndex, { value: isSampledFormula ? value : "" }];
      }),
    );
  }
}

type ColumnRun = [startColIndex: number, endColIndex: number];

// Ascending indexes, so each run of neighbours becomes one range.
function columnRuns(colIndexes: number[]): ColumnRun[] {
  return colIndexes.reduce<ColumnRun[]>((runs, colIndex) => {
    const lastRun = runs.at(-1);
    if (lastRun !== undefined && lastRun[1] === colIndex) {
      lastRun[1] = colIndex + 1;
    } else {
      runs.push([colIndex, colIndex + 1]);
    }
    return runs;
  }, []);
}

function columnRunCopies(
  source: BoundedGridRange,
  destination: BoundedGridRange,
  runs: ColumnRun[],
  pasteType: CopyPasteType,
): CopyPasteOperation[] {
  return runs.map(([startColIndex, endColIndex]) => ({
    kind: "copyPaste",
    source: columnRun(source, startColIndex, endColIndex),
    destination: columnRun(destination, startColIndex, endColIndex),
    pasteType,
  }));
}
