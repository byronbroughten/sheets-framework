import { Val } from "@byronbroughten/utils/val";

import type {
  IntendedTrait,
  PlannedColumnGroupMove,
  PlannedNewColumn,
} from "../../01_SpreadsheetSchema/configReaders/SyncPlanner";
import type { WideColumnGroupMove } from "../../01_SpreadsheetSchema/configReaders/SyncPlanner/wideAlterations";
import { beforeFirstColumnText } from "../../01_SpreadsheetSchema/configReaders/syncRunResult";
import { dimensionIds } from "../../01_SpreadsheetSchema/dimensionIds";
import type {
  ColumnMoveSpan,
  NewColumnHeadCells,
} from "../../02_SpreadsheetRaw/ClassTypes/StateRaw";
import { ColumnConfigsBaseOperator } from "../ClassBases/ColumnConfigsBaseOperator";
import { TableColumnOrder } from "./TableColumnOrder";

export interface MintedNewColumn extends PlannedNewColumn {
  tableId: string;
  columnId: string;
}

interface ColumnAt {
  tableId: string;
  colIndex: number;
}

interface ColumnInsert extends ColumnAt {
  headCells: NewColumnHeadCells;
}

// Indexed in the layout the run's column deletions and inserts leave; `table` is the codebase name a rule rewrite looks it up by.
export interface PlannedColumnMove extends ColumnMoveSpan {
  tableId: string;
  table: string;
}

export interface ColumnMovesProps {
  sent: IntendedTrait[];
  minted: MintedNewColumn[];
  groupMoves: PlannedColumnGroupMove[];
}

// Columns to sit together, in order, immediately before `before`, all by previous names.
interface OrderedColumns extends WideColumnGroupMove {
  tableId: string;
  tableKey: string;
}

interface ColumnOrderProps {
  tableId: string;
  tableKey: string;
  deleted: Set<string>;
  minted: MintedNewColumn[];
}

export class ColumnConfigsLayoutOperator extends ColumnConfigsBaseOperator {
  // Minted before any queue, so a dry run's withheld flush and the report carry the same IDs.
  mintColumnIds(newColumns: PlannedNewColumn[]): MintedNewColumn[] {
    const idPrefixesByTableId = this.tableConfigOperator.idPrefixesByTableId();
    return newColumns.map((newColumn) => {
      const tableId = this.tableConfigOperator.tableIdOf(newColumn.table);
      this._validateInsertable(tableId, newColumn);
      const idPrefix = idPrefixesByTableId.assertIdPrefix(tableId);
      return { ...newColumn, tableId, columnId: dimensionIds.col(idPrefix) };
    });
  }
  // A guard only: beforeFirstColumnStops stops the run first. Before anything is queued, so the throw sends nothing.
  private _validateInsertable(
    tableId: string,
    { table, header, before }: PlannedNewColumn,
  ): void {
    if (before === undefined) return;
    const { apiColumnIdentities } = this;
    const columnId = apiColumnIdentities.columnIdOf(table, before);
    if (!apiColumnIdentities.isFirstColumn({ tableId, columnId })) return;
    throw new Error(beforeFirstColumnText({ table, header, before }));
  }
  // At the end, a new column shifts none of the plan's other writes, so it joins their flush.
  appendColumns(minted: MintedNewColumn[]): this {
    for (const { tableId, columnId, header, before } of minted) {
      if (before !== undefined) continue;
      this.ss.raw.table(tableId).appendColumn({ columnId, header });
    }
    return this;
  }
  // Rightmost first on each Table, so a delete's index isn't moved by the ones queued before it.
  // Then the inserts, each at its index once those deletes land.
  updateDeletionsAndInserts(
    sent: IntendedTrait[],
    minted: MintedNewColumn[],
  ): this {
    const { apiColumnIdentities } = this;
    const deletions = apiColumnIdentities
      .sentColumnTraits(sent, "deleted")
      .map(({ tableId, columnId }) => ({
        tableId,
        colIndex: apiColumnIdentities.colIndexOf(tableId, columnId),
      }));
    deletions.sort((first, second) => second.colIndex - first.colIndex);
    for (const { tableId, colIndex } of deletions) {
      this.ss.raw.table(tableId).deleteColumn(colIndex);
    }
    for (const { tableId, colIndex, headCells } of this._inserts(
      minted,
      deletions,
    )) {
      this.ss.raw.table(tableId).insertColumn(colIndex, headCells);
    }
    return this;
  }
  // Rightmost first, and two before one column in reverse entry order, so the entries land in the order written.
  private _inserts(
    minted: MintedNewColumn[],
    deletions: ColumnAt[],
  ): ColumnInsert[] {
    const { apiColumnIdentities } = this;
    const inserts = minted.flatMap(
      ({ tableId, table, columnId, header, before }): ColumnInsert[] => {
        if (before === undefined) return [];
        const beforeIndex = apiColumnIdentities.colIndexOf(
          tableId,
          apiColumnIdentities.columnIdOf(table, before),
        );
        const deletedLeft = deletions.filter(
          (deletion) =>
            deletion.tableId === tableId && deletion.colIndex < beforeIndex,
        ).length;
        return [
          {
            tableId,
            colIndex: beforeIndex - deletedLeft,
            headCells: { columnId, header },
          },
        ];
      },
    );
    return inserts
      .reverse()
      .sort((first, second) => second.colIndex - first.colIndex);
  }
  // Before any queue, so a throw sends nothing; each move is indexed in the layout the run's deletions, inserts and earlier moves leave.
  planColumnMoves({
    sent,
    minted,
    groupMoves,
  }: ColumnMovesProps): PlannedColumnMove[] {
    const deleted = new Set(
      this.apiColumnIdentities
        .sentColumnTraits(sent, "deleted")
        .map(({ columnId }) => columnId),
    );
    // One order per Table, so each move is indexed in the layout the earlier ones leave.
    const orders = new Map<string, TableColumnOrder>();
    const columnGroupMoves = groupMoves.map(({ table, columns, before }) => ({
      tableId: this.tableConfigOperator.tableIdOf(table),
      tableKey: table,
      columns,
      before,
    }));
    return columnGroupMoves.flatMap((ordered) => {
      const { tableId, tableKey } = ordered;
      const order =
        orders.get(tableId) ??
        this._columnOrderAfterChanges({ tableId, tableKey, deleted, minted });
      orders.set(tableId, order);
      return this._plannedMoves(order, ordered);
    });
  }
  // Last column first, each before the one after it in the group, so the group ends in its order before its target.
  private _plannedMoves(
    order: TableColumnOrder,
    { tableId, tableKey, columns, before }: OrderedColumns,
  ): PlannedColumnMove[] {
    const { apiColumnIdentities } = this;
    const sequence = [...columns, before];
    return columns
      .map((column, index) => ({
        column,
        before: Val.assert(sequence[index + 1], `column after ${column}`),
      }))
      .reverse()
      .flatMap(({ column, before: target }): PlannedColumnMove[] => {
        const span = order.moveSpan({
          columnId: apiColumnIdentities.columnIdOf(tableKey, column),
          beforeColumnId: apiColumnIdentities.columnIdOf(tableKey, target),
          column,
          before: target,
        });
        if (span === undefined) return [];
        order.move(span);
        return [{ tableId, table: tableKey, ...span }];
      });
  }
  // Queued once the run's deletions and inserts have flushed, since their indexes are the layout those leave.
  moveColumns(moves: PlannedColumnMove[]): this {
    for (const { tableId, colIndex, count, beforeColIndex } of moves) {
      this.ss.raw
        .table(tableId)
        .moveColumns({ colIndex, count, beforeColIndex });
    }
    return this;
  }
  // As updateDeletionsAndInserts and appendColumns leave the Table: each insert before its column in entry order, each append at the end.
  private _columnOrderAfterChanges({
    tableId,
    tableKey,
    deleted,
    minted,
  }: ColumnOrderProps): TableColumnOrder {
    const order = this._columnOrder(tableId, tableKey).without(deleted);
    for (const {
      tableId: mintedTableId,
      table: key,
      columnId,
      before,
    } of minted) {
      if (mintedTableId !== tableId) continue;
      let beforeColumnId: string | undefined;
      if (before !== undefined) {
        beforeColumnId = this.apiColumnIdentities.columnIdOf(key, before);
      }
      order.insert(columnId, beforeColumnId);
    }
    return order;
  }
  // Public for ColumnConfigsOperator's traitReadings, which reads each column's `before` from it.
  _columnOrder(tableId: string, tableKey: string): TableColumnOrder {
    const table = this.ss.raw.table(tableId);
    return new TableColumnOrder(
      tableKey,
      table.fullTableColIndexes.map((colIndex) =>
        table.columnResolver.columnIdAt(colIndex),
      ),
    );
  }
}
