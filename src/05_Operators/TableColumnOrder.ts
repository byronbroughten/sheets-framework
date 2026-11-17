import { Arr, type BlockMove } from "@byronbroughten/utils/arr";

import type { ColumnMoveSpan } from "../02_SpreadsheetRaw/ClassTypes/StateRaw";

export interface ColumnOrderMove {
  columnId: string;
  beforeColumnId: string;
  column: string;
  before: string;
}

// A Table's column IDs in sheet order, as ColumnConfigsOperator plans a run's column moves on them.
export class TableColumnOrder {
  private columnIds: (string | undefined)[];
  // A column with no ID yet, which the Raw tier reads as "", keeps its place so the indexes still count it.
  constructor(
    readonly tableKey: string,
    liveColumnIds: string[],
  ) {
    this.columnIds = liveColumnIds.map((columnId) =>
      columnId === "" ? undefined : columnId,
    );
  }
  columnIdAfter(columnId: string): string | undefined {
    return this.columnIds[this.columnIds.indexOf(columnId) + 1];
  }
  without(deleted: ReadonlySet<string>): this {
    this.columnIds = this.columnIds.filter(
      (columnId) => columnId === undefined || !deleted.has(columnId),
    );
    return this;
  }
  // As TableRaw.insertColumn leaves it before the named column, or appendColumn at the end.
  insert(columnId: string, beforeColumnId: string | undefined): this {
    if (beforeColumnId === undefined) {
      this.columnIds.push(columnId);
      return this;
    }
    this.columnIds.splice(this.columnIds.indexOf(beforeColumnId), 0, columnId);
    return this;
  }
  // Undefined when the order already has the column there.
  moveSpan({
    columnId,
    beforeColumnId,
    column,
    before,
  }: ColumnOrderMove): ColumnMoveSpan | undefined {
    const colIndex = this.columnIds.indexOf(columnId);
    const beforeColIndex = this.columnIds.indexOf(beforeColumnId);
    if (
      colIndex === -1 ||
      beforeColIndex === -1 ||
      colIndex === beforeColIndex
    ) {
      throw new Error(
        `Table ${this.tableKey} can't move ${column} before ${before}: the run deletes one of them, or they are the same column.`,
      );
    }
    if (beforeColIndex === colIndex + 1) return undefined;
    if (beforeColIndex !== 0) return { colIndex, count: 1, beforeColIndex };
    if (colIndex === this.columnIds.length - 1) {
      throw new Error(
        `Table ${this.tableKey} can't move ${column} before ${before}, its first column: moving its last column to the front is unmeasured.`,
      );
    }
    // A move to the first column lands outside the Table, so the columns before it move right instead (sheets-behaviour.md).
    return { colIndex: 0, count: colIndex, beforeColIndex: colIndex + 1 };
  }
  move(span: ColumnMoveSpan): this {
    this.columnIds = Arr.moveBlock(this.columnIds, blockMoveOf(span));
    return this;
  }
}

export function blockMoveOf({
  colIndex,
  count,
  beforeColIndex,
}: ColumnMoveSpan): BlockMove {
  return { index: colIndex, count, beforeIndex: beforeColIndex };
}
