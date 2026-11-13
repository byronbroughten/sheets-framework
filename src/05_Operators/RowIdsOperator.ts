import {
  fillRowIdsName,
  type SyncNotice,
  type SyncPlanItem,
} from "../01_SpreadsheetSchema/configReaders/syncRunResult";
import { dimensionIds } from "../01_SpreadsheetSchema/dimensionIds";
import type { TableRaw } from "../02_SpreadsheetRaw/TableRaw";
import { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import { SpreadsheetBaseOperator } from "./SpreadsheetBaseOperator";
import { TableConfigsOperator } from "./TableConfigsOperator";

interface IdColumn {
  table: TableRaw;
  tableKey: string;
  colIndex: number;
}

interface BlankIdColumn extends IdColumn {
  idPrefix: string;
  rowIndexes: number[];
}

// Not SpreadsheetNamed.fillMissingRowIds: that reads the installed configs, and sync fills from the ones it just regenerated.
export class RowIdsOperator extends SpreadsheetBaseOperator {
  get ss(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  get tableConfigOperator(): TableConfigsOperator {
    return new TableConfigsOperator(this.operatorProps);
  }
  // From the pre-send read, so the stop comes before any send.
  validateIdPrefixes(): void {
    const idPrefixesByTableId = this.tableConfigOperator.idPrefixesByTableId();
    for (const { table } of this._liveIdColumns()) {
      if (idPrefixesByTableId.get(table.tableId)) continue;
      throw new Error(
        `Table "${table.name}" has an ID column but no ID prefix, so sync stopped before sending.`,
      );
    }
  }
  // Not a claim (#286): a cell holding anything is left alone.
  checkboxPlan(): SyncPlanItem[] {
    return this._columnsLackingCheckbox().map(({ tableKey }) =>
      idPlanItem(tableKey, `add the ${fillRowIdsName} checkbox`),
    );
  }
  addCheckboxes(): this {
    for (const { table, colIndex } of this._columnsLackingCheckbox()) {
      table
        .column(colIndex)
        .headCell("action")
        .updateValue(false)
        .addCheckboxValidation();
    }
    return this;
  }
  // The Table name chip covers the action-row cell above a Table's first column.
  idFirstNotices(): SyncNotice[] {
    return this._liveIdColumns().flatMap(({ tableKey, colIndex }) => {
      if (colIndex !== 0) return [];
      return [{ kind: "idFirst", table: tableKey }];
    });
  }
  // Planned from the fresh configs.
  fillPlan(): SyncPlanItem[] {
    return this._blankIdColumns().map((column) => fillPlanItem(column));
  }
  // Queues the minted IDs unflushed, so the caller's flush is the fill's own send.
  fillBlankIds(): SyncPlanItem[] {
    const items: SyncPlanItem[] = [];
    for (const column of this._blankIdColumns()) {
      items.push(fillPlanItem(column, mintRowIds(column)));
    }
    return items;
  }
  private _blankIdColumns(): BlankIdColumn[] {
    const idColumns = this._freshIdColumns();
    for (const { table, colIndex } of idColumns) {
      table.column(colIndex).gatherFetchFull();
    }
    this.ss.identified.fetchAllGathered(true);
    return idColumns.flatMap((idColumn): BlankIdColumn[] => {
      const column = idColumn.table.column(idColumn.colIndex);
      const rowIndexes = column.cellIndexesFull.filter(
        (rowIndex) => column.valueOrEmpty(rowIndex) === "",
      );
      if (rowIndexes.length === 0) return [];
      return [{ ...idColumn, rowIndexes }];
    });
  }
  private _freshIdColumns(): (IdColumn & { idPrefix: string })[] {
    const { raw } = this.ss;
    return Object.entries(this.tableConfigOperator.newTableConfigs()).flatMap(
      ([tableKey, { tableId, idPrefix, hasIdColumn }]) => {
        if (!hasIdColumn) return [];
        const idColumn = idColumnOf(raw.table(tableId), tableKey);
        if (idColumn === undefined) {
          throw new Error(`Table "${tableKey}" has no ID column to fill.`);
        }
        return [{ ...idColumn, idPrefix }];
      },
    );
  }
  private _columnsLackingCheckbox(): IdColumn[] {
    return this._liveIdColumns().filter(
      ({ table, colIndex }) =>
        table.column(colIndex).headCell("action").valueOrEmpty() === "",
    );
  }
  private _liveIdColumns(): IdColumn[] {
    const { raw } = this.ss;
    return this.tableConfigOperator
      .tableIdsApiAccesses()
      .flatMap((tableId): IdColumn[] => {
        const table = raw.table(tableId);
        const idColumn = idColumnOf(table, this.schema.titleToName(table.name));
        if (idColumn === undefined) return [];
        return [idColumn];
      });
  }
}

const idColumnName = "id";

function idPlanItem(tableKey: string, what: string): SyncPlanItem {
  return {
    subject: { table: tableKey, column: idColumnName },
    what,
    effects: [],
  };
}

function fillPlanItem(
  { tableKey, rowIndexes }: BlankIdColumn,
  rowIds?: string[],
): SyncPlanItem {
  const count = `${rowIndexes.length} blank ${rowIndexes.length === 1 ? "ID" : "IDs"}`;
  const listed = rowIds === undefined ? "" : `: ${rowIds.join(", ")}`;
  return idPlanItem(tableKey, `fill ${count}${listed}`);
}

function mintRowIds({
  table,
  colIndex,
  idPrefix,
  rowIndexes,
}: BlankIdColumn): string[] {
  const column = table.column(colIndex);
  const rowIds: string[] = [];
  for (const rowIndex of rowIndexes) {
    const rowId = dimensionIds.row(idPrefix);
    column.updateValue(rowIndex, rowId);
    rowIds.push(rowId);
  }
  return rowIds;
}

function idColumnOf(table: TableRaw, tableKey: string): IdColumn | undefined {
  const colIndex = table.fullTableColIndexes.find(
    (index) => table.column(index).profile.isIdColumn,
  );
  if (colIndex === undefined) return undefined;
  return { table, tableKey, colIndex };
}
