import type {
  BoundedGridRange,
  InsertRangeOperation,
} from "../../00_Source/RawSource/RawSource";
import {
  hasQueuedWrites,
  TableBaseRaw,
  tableColumnLabel,
} from "../ClassBases/TableBaseRaw";
import { columnCountChange } from "../ClassBases/TableCommonRaw";
import type {
  ChangeTableColumns,
  ColumnMoveSpan,
  NewColumnHeadCells,
  TableColumnChange,
  TableColumnChangeOperation,
} from "../ClassTypes/StateRaw";
import { TableRaw } from "../TableRaw";

const columnChangeRequests = {
  insertColumn: { kind: "insertRange", action: "Inserting a column inside" },
  deleteColumn: { kind: "deleteRange", action: "Deleting a column of" },
  moveColumns: { kind: "cutPaste", action: "Moving a column of" },
} as const satisfies Record<
  TableColumnChange["kind"],
  { kind: TableColumnChangeOperation["kind"]; action: string }
>;

type ColumnCountChange = Exclude<TableColumnChange, { kind: "moveColumns" }>;
type ColumnMove = Extract<TableColumnChange, { kind: "moveColumns" }>;

export class TableColumnQueuerRaw extends TableBaseRaw {
  get table(): TableRaw {
    return new TableRaw(this.tableRawProps);
  }
  // Past the inserts already queued, since each lands at the Table end as it stands then.
  appendColumn(headCells: NewColumnHeadCells): number {
    const { table } = this;
    const colIndex = table.columnCount + table.writes.insertTableEndColumnCount;
    table.queueTableWrite({ action: "insertTableEndColumn" });
    table.column(colIndex)._initHeadCells(headCells);
    return colIndex;
  }
  // Strictly inside, so the band grows the Table with no range update; appendColumn adds at the end.
  insertColumn(colIndex: number, headCells: NewColumnHeadCells): void {
    this.table.assertRowIndexesNotStale();
    this._assertStrictlyInside(
      colIndex,
      `a column inserted at ${tableColumnLabel(colIndex)}`,
    );
    this.table.queueTableWrite({
      action: "insertColumn",
      colIndex,
      headCells,
    });
  }
  deleteColumn(colIndex: number): void {
    const { table } = this;
    table.assertRowIndexesNotStale();
    const columnCount = this._columnCountAfterChanges();
    if (
      !Number.isInteger(colIndex) ||
      colIndex < 0 ||
      colIndex >= columnCount
    ) {
      throw new Error(
        `${table.tableLabel} has no ${tableColumnLabel(colIndex)} to delete: it has ${columnCount} columns.`,
      );
    }
    if (columnCount === 1) {
      throw new Error(
        `Refusing to delete the last column of ${table.tableLabel}: a Table keeps at least one.`,
      );
    }
    table.queueTableWrite({ action: "deleteColumn", colIndex });
  }
  // Lands strictly inside, like insertColumn: an insert at the Table's first column lands outside it (sheets-behaviour.md).
  moveColumns({ colIndex, count, beforeColIndex }: ColumnMoveSpan): void {
    const { table } = this;
    table.assertRowIndexesNotStale();
    const columnCount = this._columnCountAfterChanges();
    const isBlockInside =
      Number.isInteger(colIndex) &&
      Number.isInteger(count) &&
      count >= 1 &&
      colIndex >= 0 &&
      colIndex + count <= columnCount;
    if (!isBlockInside) {
      throw new Error(
        `${table.tableLabel} has no ${count} column(s) from ${tableColumnLabel(colIndex)} to move: it has ${columnCount} columns.`,
      );
    }
    this._assertStrictlyInside(
      beforeColIndex,
      `columns moved before ${tableColumnLabel(beforeColIndex)}`,
    );
    if (colIndex <= beforeColIndex && beforeColIndex <= colIndex + count) {
      throw new Error(
        `${table.tableLabel} already has its ${count} column(s) from ${tableColumnLabel(colIndex)} before ${tableColumnLabel(beforeColIndex)}.`,
      );
    }
    table.queueTableWrite({
      action: "moveColumns",
      colIndex,
      count,
      beforeColIndex,
    });
  }
  get columnInsertLabel(): string {
    return `Inserting a column at the end of ${this._namedTableOnSheet}`;
  }
  columnInsertSplitting(neighbourName: string): string {
    return splittingPhrase(this.columnInsertLabel, neighbourName);
  }
  // A move's gap is the insert its cutPaste follows, as _columnMoveOperations orders them.
  columnChangeSplitting(
    { operations }: ChangeTableColumns,
    insert: InsertRangeOperation,
    neighbourName: string,
  ): string {
    const next = operations[operations.indexOf(insert) + 1];
    const changeKind =
      next?.kind === "cutPaste" ? "moveColumns" : "insertColumn";
    return splittingPhrase(this._columnChangeLabel(changeKind), neighbourName);
  }
  // From the top head row down, so the head rows move with the Table; the widen is what grows it.
  gatherInsertTableEndColumnsOperation(): void {
    const { table } = this;
    const insertCount = table.writes.insertTableEndColumnCount;
    if (insertCount === 0) return;
    const { columnCount } = table;
    const newColumns = this._headAndBodyBand(
      columnCount,
      columnCount + insertCount,
    );
    this._validateNoNeighbourSplitBy(newColumns, this.columnInsertLabel);
    table.sheet.queueGridColumnsThrough(newColumns.endColumnIndex);
    table.writeOperations.insertTableEndColumns.push({
      tableId: table.tableId,
      newColumns,
      operations: [
        { kind: "insertRange", range: newColumns, shiftDimension: "COLUMNS" },
        {
          kind: "updateTableRange",
          tableId: table.tableId,
          range: {
            ...newColumns,
            startRowIndex: table.startRowIndex,
            startColumnIndex: table.startColumnIndex,
          },
        },
      ],
    });
    table.growColumnCount(insertCount);
  }
  // Sent in a flush of its own on this Table, since every other queued write is indexed by the columns before it.
  gatherColumnChangesOperation(): void {
    const { table } = this;
    const { columnChanges } = table.writes;
    if (columnChanges.length === 0) return;
    this._assertColumnChangesAlone();
    const startColumnIndex = table.originAtGathering().sheetColIndex(0);
    const operations = columnChanges.flatMap((change) => {
      if (change.kind === "moveColumns") {
        return this._columnMoveOperations(change);
      }
      return [this._columnChangeOperation(change)];
    });
    table.markColumnPropertiesStale();
    table.writeOperations.changeTableColumns.push({
      tableId: table.tableId,
      sheetId: this.sheetGid,
      startColumnIndex,
      operations,
    });
  }
  private _assertStrictlyInside(colIndex: number, taken: string): void {
    const columnCount = this._columnCountAfterChanges();
    if (Number.isInteger(colIndex) && colIndex >= 1 && colIndex < columnCount) {
      return;
    }
    throw new Error(
      `${this.table.tableLabel} can't take ${taken}: it must sit strictly inside its ${columnCount} columns, so from 1 to ${columnCount - 1}.`,
    );
  }
  private _columnCountAfterChanges(): number {
    const { table } = this;
    return table.writes.columnChanges.reduce(
      (count, { kind }) =>
        count + columnCountChange[columnChangeRequests[kind].kind],
      table.columnCount,
    );
  }
  private _validateNoNeighbourSplitBy(
    band: BoundedGridRange,
    actionLabel: string,
  ): void {
    const split = this.table.neighbourTables.find((neighbour) =>
      neighbour.isSplitBy(band),
    );
    if (split === undefined) return;
    throw new Error(
      `${splittingPhrase(actionLabel, split.name)} with its head rows. Move "${split.name}" so that it and its head rows sit entirely within rows ${band.startRowIndex + 1}–${band.endRowIndex}, or entirely outside them.`,
    );
  }
  private _assertColumnChangesAlone(): void {
    const { table } = this;
    const otherWrites = {
      ...this.tableState.writeQueue,
      table: { ...table.writes, columnChanges: [] },
    };
    if (!hasQueuedWrites(otherWrites)) return;
    throw new Error(
      `Refusing to change the columns of ${table.tableLabel} alongside its other queued writes: they are indexed by the columns before the change. Flush the column change on its own.`,
    );
  }
  // Each change is read in the layout the ones before it leave, so the working view moves with it.
  private _columnChangeOperation(
    change: ColumnCountChange,
  ): TableColumnChangeOperation {
    const { colIndex } = change;
    const { kind } = columnChangeRequests[change.kind];
    const band = this._headAndBodyBand(colIndex, colIndex + 1);
    this._validateNoNeighbourSplitBy(
      band,
      this._columnChangeLabel(change.kind),
    );
    const shiftCount = columnCountChange[kind];
    this._shiftWorkingColumns(colIndex, shiftCount);
    this.table.growColumnCount(shiftCount);
    if (change.kind === "insertColumn") {
      this.table.column(colIndex)._initHeadCells(change.headCells);
    }
    return { kind, range: band, shiftDimension: "COLUMNS" };
  }
  // Not moveDimension, which reorders a Table stacked in the same sheet columns (sheets-behaviour.md).
  private _columnMoveOperations(
    move: ColumnMove,
  ): TableColumnChangeOperation[] {
    const { colIndex, count, beforeColIndex } = move;
    const gap = this._headAndBodyBand(beforeColIndex, beforeColIndex + count);
    this._validateNoNeighbourSplitBy(
      gap,
      this._columnChangeLabel("moveColumns"),
    );
    // An insert running past the grid's edge is clipped to it (sheets-behaviour.md), losing a cut column.
    this.table.sheet.queueGridColumnsThrough(gap.endColumnIndex);
    const colIndexAfterGap =
      colIndex > beforeColIndex ? colIndex + count : colIndex;
    const vacated = this._headAndBodyBand(
      colIndexAfterGap,
      colIndexAfterGap + count,
    );
    this._rekeyWorkingColumns((byColIndex) => movedColumns(byColIndex, move));
    return [
      { kind: "insertRange", range: gap, shiftDimension: "COLUMNS" },
      {
        kind: "cutPaste",
        source: vacated,
        destination: {
          sheetId: gap.sheetId,
          rowIndex: gap.startRowIndex,
          columnIndex: gap.startColumnIndex,
        },
      },
      { kind: "deleteRange", range: vacated, shiftDimension: "COLUMNS" },
    ];
  }
  // From the top head row down, so the head rows move with the Table.
  private _headAndBodyBand(
    startColIndex: number,
    endColIndex: number,
  ): BoundedGridRange {
    const { table } = this;
    const origin = table.originAtGathering();
    return {
      sheetId: this.sheetGid,
      startRowIndex: origin.topHeadSheetRowIndex,
      endRowIndex: origin.sheetRowIndex(table.dataRowCount),
      startColumnIndex: origin.sheetColIndex(startColIndex),
      endColumnIndex: origin.sheetColIndex(endColIndex),
    };
  }
  private _shiftWorkingColumns(colIndex: number, shiftCount: 1 | -1): void {
    this._rekeyWorkingColumns((byColIndex) =>
      shiftedColumns(byColIndex, colIndex, shiftCount),
    );
    if (shiftCount !== 1) return;
    for (const rowState of this.rowStates.values()) {
      rowState.set(colIndex, { value: "" });
    }
  }
  private _rekeyWorkingColumns(
    rekey: <VL>(byColIndex: Map<number, VL>) => Map<number, VL>,
  ): void {
    for (const [rowIndex, rowState] of this.rowStates) {
      this.rowStates.set(rowIndex, rekey(rowState));
    }
    this.tableState.working.columnStates = rekey(this.columnStates);
    // An earlier change's new head cells are queued by colIndex, so a later change in the same gather carries them.
    for (const rowWrites of this.tableState.writeQueue.rows.values()) {
      rowWrites.fillCells = rekey(rowWrites.fillCells);
    }
  }
  private get _namedTableOnSheet(): string {
    return `Table "${this.table.name}" on sheet "${this.sheetTitle}"`;
  }
  private _columnChangeLabel(changeKind: TableColumnChange["kind"]): string {
    return `${columnChangeRequests[changeKind].action} ${this._namedTableOnSheet}`;
  }
}

function splittingPhrase(actionLabel: string, neighbourName: string): string {
  return `${actionLabel} would shift only part of Table "${neighbourName}"`;
}

function shiftedColumns<VL>(
  byColIndex: Map<number, VL>,
  colIndex: number,
  shiftCount: 1 | -1,
): Map<number, VL> {
  return remappedColumns(byColIndex, (index) => {
    if (index < colIndex) return index;
    if (index > colIndex || shiftCount === 1) return index + shiftCount;
    return undefined;
  });
}

function movedColumns<VL>(
  byColIndex: Map<number, VL>,
  span: ColumnMoveSpan,
): Map<number, VL> {
  return remappedColumns(byColIndex, (index) => movedColIndex(index, span));
}

// A column the remap gives no index goes.
function remappedColumns<VL>(
  byColIndex: Map<number, VL>,
  newColIndex: (index: number) => number | undefined,
): Map<number, VL> {
  return Array.from(byColIndex).reduce((remapped, [index, value]) => {
    const remappedIndex = newColIndex(index);
    if (remappedIndex !== undefined) remapped.set(remappedIndex, value);
    return remapped;
  }, new Map<number, VL>());
}

// Where a column sits once the block has moved; the columns it passes close the gap.
function movedColIndex(
  index: number,
  { colIndex, count, beforeColIndex }: ColumnMoveSpan,
): number {
  const isInBlock = colIndex <= index && index < colIndex + count;
  if (beforeColIndex > colIndex) {
    if (isInBlock) return index + beforeColIndex - colIndex - count;
    if (colIndex + count <= index && index < beforeColIndex) {
      return index - count;
    }
    return index;
  }
  if (isInBlock) return index - (colIndex - beforeColIndex);
  if (beforeColIndex <= index && index < colIndex) return index + count;
  return index;
}
