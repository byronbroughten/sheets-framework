import { Val } from "@byronbroughten/utils/val";

import type {
  BoundedGridRange,
  TableSnapshot,
} from "../../00_Source/RawSource/RawSource";
import { TableOrigin } from "../../01_SpreadsheetSchema/TableOrigin";
import { emptyStateRaw } from "../ClassTypes/emptyStateRaw";
import type {
  BoundedTableGridRange,
  ColumnStateRaw,
  ColumnStatesRaw,
  RowStateRaw,
  RowStatesRaw,
  TablePropertiesRaw,
  TablesStateRaw,
  TableStateRaw,
  TableWriteQueueRaw,
} from "../ClassTypes/StateRaw";
import { SheetBaseRaw, type SheetRawProps } from "./SheetBaseRaw";

export interface TableAddressRaw {
  tableId: string;
}
// `sheetGid` places a Table not yet in state, and must match a fetched one's.
export interface TableRawProps extends SheetRawProps, TableAddressRaw {}

export class TableBaseRaw extends SheetBaseRaw {
  private readonly addressedTableId: string;
  constructor({ tableId, ...props }: TableRawProps) {
    super(props);
    this.addressedTableId = tableId;
    validateTableOnSheet(this.tablesStateRaw, tableId, this.sheetGid);
    this._ensureSheetState();
    this._ensureTableBeforePropertiesById();
  }
  private _ensureTableBeforePropertiesById(): void {
    const tableId = this.addressedTableId;
    const { tablesBeforePropertiesById } = this.sheetState;
    if (this.tablesStateRaw.has(tableId)) return;
    if (tablesBeforePropertiesById.has(tableId)) return;
    tablesBeforePropertiesById.set(
      tableId,
      emptyStateRaw.tableState(this.sheetGid),
    );
  }
  protected get tableState(): TableStateRaw {
    const tableId = this.addressedTableId;
    return Val.assert(
      this.tablesStateRaw.get(tableId) ??
        this.sheetState.tablesBeforePropertiesById.get(tableId),
      `Table state for tableId ${tableId}`,
    );
  }
  // Absent until fetched.
  protected get tableProperties(): TablePropertiesRaw | undefined {
    return this.tableState.properties;
  }
  get columnStates(): ColumnStatesRaw {
    return this.tableState.working.columnStates;
  }
  get rowStates(): RowStatesRaw {
    return this.tableState.working.rowStates;
  }
  get sheetTitle(): string {
    return this.sheetState.working.title ?? untitledSheetTitle;
  }
  get sheetLabel(): string {
    return sheetLabel(this.sheetState.working.title, this.sheetGid);
  }
  get tableRawProps(): TableRawProps {
    return {
      tableId: this.addressedTableId,
      sheetGid: this.sheetGid,
      ...this.spreadsheetRawProps,
    };
  }
  tableIds(): string[] {
    return tableIdsOnSheet(this.tablesStateRaw, this.sheetGid);
  }
  hasOneTable(): boolean {
    return this.tableIds().length === 1;
  }
  getRowState(rowIndex: number): RowStateRaw {
    return Val.assert(
      this.rowStates.get(rowIndex),
      `rowState for ${this.rowLabel(rowIndex)} on sheetGid ${this.sheetGid}`,
    );
  }
  // The live Table once fetched; before that, the spot the framework creates Tables at.
  tableOrigin(): TableOrigin {
    const properties = this.tableProperties;
    if (properties === undefined) return TableOrigin.expected();
    return originOf(properties);
  }
  rowLabel(rowIndex: number): string {
    return `row ${this.tableOrigin().rowNumber(rowIndex)}`;
  }
  protected _boundedSheetGridRange(
    {
      startRowIndex,
      endRowIndex,
      startColIndex,
      endColIndex,
    }: BoundedTableGridRange,
    origin: TableOrigin,
  ): BoundedGridRange {
    return {
      sheetId: this.sheetGid,
      startRowIndex: origin.sheetRowIndex(startRowIndex),
      endRowIndex: origin.sheetRowIndex(endRowIndex),
      startColumnIndex: origin.sheetColIndex(startColIndex),
      endColumnIndex: origin.sheetColIndex(endColIndex),
    };
  }
  protected _ensureColumnState(colIndex: number): ColumnStateRaw {
    return ensureColumnState(this.tableState, colIndex);
  }
}

const untitledSheetTitle = "(untitled)";

export function sheetLabel(
  title: string | undefined,
  sheetGid: number,
): string {
  return `"${title ?? untitledSheetTitle}" (gid ${sheetGid})`;
}

export function tableOnSheetLabel({
  name,
  onSheet,
}: {
  name: string;
  onSheet: string;
}): string {
  return `Table "${name}" on ${onSheet}`;
}

export function tableIdsOnSheet(
  tables: TablesStateRaw,
  sheetGid: number,
): string[] {
  return Array.from(tables.entries())
    .filter(([, tableState]) => tableState.sheetGid === sheetGid)
    .map(([tableId]) => tableId);
}

// A Table reached through a sheet it isn't on would read and write the wrong grid.
function validateTableOnSheet(
  tables: TablesStateRaw,
  tableId: string,
  sheetGid: number,
): void {
  const liveSheetGid = tables.get(tableId)?.sheetGid;
  if (liveSheetGid === undefined || liveSheetGid === sheetGid) return;
  throw new Error(
    `Table ${tableId} is on sheetGid ${liveSheetGid}, not sheetGid ${sheetGid}, which reached it.`,
  );
}

export function tableStateOf(
  tables: TablesStateRaw,
  tableId: string,
): TableStateRaw {
  return Val.assert(tables.get(tableId), `Table state for tableId ${tableId}`);
}

export function hasQueuedWrites({ table, rows }: TableWriteQueueRaw): boolean {
  return (
    rows.size > 0 ||
    table.sort !== undefined ||
    table.insertTableEndColumnCount > 0 ||
    table.fillColumns.length > 0 ||
    table.findReplaces.length > 0 ||
    table.columnTypes.size > 0 ||
    table.checkboxCells.length > 0 ||
    table.conditionalFormatRules.length > 0 ||
    table.editProtections.length > 0
  );
}

export function originOf(properties: TablePropertiesRaw): TableOrigin {
  return new TableOrigin({
    headerRowIndex: properties.startRowIndex,
    startColIndex: properties.startColumnIndex,
  });
}

export function integrateColumnProperties(
  tableState: TableStateRaw,
  table: TableSnapshot,
): void {
  tableState.working.columnStates.forEach((columnState) => {
    delete columnState.validationValues;
    delete columnState.validationConditionType;
    delete columnState.columnType;
  });
  table.columnProperties.forEach((colProps) => {
    // The API states columnIndex table-relative, as Raw does.
    const columnState = ensureColumnState(tableState, colProps.columnIndex);
    if (colProps.dataValidationValues.length > 0) {
      columnState.validationValues = colProps.dataValidationValues;
    }
    if (colProps.dataValidationConditionType !== undefined) {
      columnState.validationConditionType =
        colProps.dataValidationConditionType;
    }
    if (colProps.columnType !== undefined) {
      columnState.columnType = colProps.columnType;
    }
  });
  // Queued column types outlive a re-fetch until the flush sends them.
  tableState.writeQueue.table.columnTypes.forEach((columnType, colIndex) => {
    ensureColumnState(tableState, colIndex).columnType = columnType;
  });
}

function ensureColumnState(
  tableState: TableStateRaw,
  colIndex: number,
): ColumnStateRaw {
  const columnStates = tableState.working.columnStates;
  const existing = columnStates.get(colIndex);
  if (existing !== undefined) return existing;
  const created: ColumnStateRaw = {};
  columnStates.set(colIndex, created);
  return created;
}
