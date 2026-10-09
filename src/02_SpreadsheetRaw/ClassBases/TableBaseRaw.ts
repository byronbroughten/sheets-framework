import { Val } from "@byronbroughten/utils/val";

import type { TableSnapshot } from "../../00_Source/RawSource/RawSource";
import { SpreadsheetSchema } from "../../01_SpreadsheetSchema/configReaders/SpreadsheetSchema";
import { tableConfigsByTableId } from "../../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { TableOrigin } from "../../01_SpreadsheetSchema/TableOrigin";
import { emptyStateRaw } from "../ClassTypes/emptyStateRaw";
import type {
  ColumnStateRaw,
  ColumnStatesRaw,
  RowStateRaw,
  RowStatesRaw,
  SheetStateRaw,
  TablePropertiesRaw,
  TablesStateRaw,
  TableStateRaw,
  TableWriteQueueRaw,
} from "../ClassTypes/StateRaw";
import {
  SpreadsheetBaseRaw,
  type SpreadsheetRawProps,
} from "./SpreadsheetBaseRaw";

// `ss.tableOnSheet(gid)` reaches a Table through its sheet: its only one, else the one its configs record.
export type TableAddressRaw = { sheetGid: number } | { tableId: string };
export type TableRawProps = SpreadsheetRawProps & TableAddressRaw;

export class TableBaseRaw extends SpreadsheetBaseRaw {
  readonly sheetGid: number;
  private readonly tableAddress: TableAddressRaw;
  constructor({ spreadsheetStateRaw, ...tableAddress }: TableRawProps) {
    super({ spreadsheetStateRaw });
    this.tableAddress = tableAddress;
    this.sheetGid = sheetGidOf(spreadsheetStateRaw.tables, tableAddress);
    this._ensureSheetState();
    this._ensureTableBeforePropertiesById();
  }
  private _ensureSheetState(): void {
    if (!this.sheetsStateRaw.has(this.sheetGid)) {
      this.sheetsStateRaw.set(
        this.sheetGid,
        emptyStateRaw.sheetState(this.sheetGid),
      );
    }
  }
  private _ensureTableBeforePropertiesById(): void {
    if (!("tableId" in this.tableAddress)) return;
    const { tableId } = this.tableAddress;
    const { tablesBeforePropertiesById } = this.sheetState;
    if (this.tablesStateRaw.has(tableId)) return;
    if (tablesBeforePropertiesById.has(tableId)) return;
    tablesBeforePropertiesById.set(
      tableId,
      emptyStateRaw.tableState(this.sheetGid),
    );
  }
  protected get sheetState(): SheetStateRaw {
    return Val.assert(
      this.sheetsStateRaw.get(this.sheetGid),
      `sheetState for sheetGid ${this.sheetGid}`,
    );
  }
  protected get tableState(): TableStateRaw {
    return this._resolveTableState();
  }
  private _resolveTableState(): TableStateRaw {
    if ("tableId" in this.tableAddress) {
      const { tableId } = this.tableAddress;
      return Val.assert(
        this.tablesStateRaw.get(tableId) ??
          this.sheetState.tablesBeforePropertiesById.get(tableId),
        `Table state for tableId ${tableId}`,
      );
    }
    const tableId = this.tableIdReachedByGid();
    if (tableId === undefined) return this.sheetState.tableBeforeProperties;
    return tableStateOf(this.tablesStateRaw, tableId);
  }
  // Absent until fetched, and for a sheet that holds no Table or several.
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
    return this.sheetState.working.title ?? "(untitled)";
  }
  get sheetLabel(): string {
    return sheetLabel(this.sheetState.working.title, this.sheetGid);
  }
  get tableRawProps(): TableRawProps {
    return {
      ...this.tableAddress,
      ...this.spreadsheetRawProps,
    };
  }
  tableIds(): string[] {
    return Array.from(this.tablesStateRaw.entries())
      .filter(([, tableState]) => tableState.sheetGid === this.sheetGid)
      .map(([tableId]) => tableId);
  }
  // Absent for a sheet that holds no Table or several.
  onlyTableId(): string | undefined {
    const [tableId, ...otherTableIds] = this.tableIds();
    if (otherTableIds.length > 0) return undefined;
    return tableId;
  }
  tableIdReachedByGid(): string | undefined {
    return tableIdReachedAmong(this.tableIds(), this.sheetGid);
  }
  hasOneTable(): boolean {
    return this.onlyTableId() !== undefined;
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
  protected _updateWorkingTableName(tableId: string, name: string): void {
    const properties = this.tablesStateRaw.get(tableId)?.properties;
    if (properties !== undefined) properties.name = name;
  }
  protected _ensureColumnState(colIndex: number): ColumnStateRaw {
    return ensureColumnState(this.tableState, colIndex);
  }
}

export function sheetLabel(
  title: string | undefined,
  sheetGid: number,
): string {
  return `"${title ?? "(untitled)"}" (gid ${sheetGid})`;
}

// A recorded Table is reachable by its ID before its properties are fetched.
function sheetGidOf(
  tables: TablesStateRaw,
  tableAddress: TableAddressRaw,
): number {
  if ("sheetGid" in tableAddress) return tableAddress.sheetGid;
  const { tableId } = tableAddress;
  return Val.assert(
    tables.get(tableId) ?? tableConfigsByTableId().get(tableId),
    `Table state for tableId ${tableId}`,
  ).sheetGid;
}

// The sheet's only Table, else the one its configs record.
export function tableIdReachedAmong(
  tableIds: string[],
  sheetGid: number,
): string | undefined {
  const [tableId, ...otherTableIds] = tableIds;
  if (otherTableIds.length === 0) return tableId;
  const recordedTableIds = new Set(
    new SpreadsheetSchema().tablesOnGid(sheetGid).map((table) => table.tableId),
  );
  const [recordedTableId, ...otherRecordedTableIds] = tableIds.filter((id) =>
    recordedTableIds.has(id),
  );
  if (otherRecordedTableIds.length > 0) return undefined;
  return recordedTableId;
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
