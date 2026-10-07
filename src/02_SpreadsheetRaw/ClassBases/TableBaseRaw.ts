import type {
  SheetSnapshot,
  TableSnapshot,
} from "../../00_Source/RawSource/RawSource";
import { SpreadsheetSchema } from "../../01_SpreadsheetSchema/SpreadsheetSchema";
import { TableOrigin } from "../../01_SpreadsheetSchema/TableOrigin";
import { Val } from "../../utils/Val";
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

// `ss.tableOnSheet(gid)` reaches a Table through its sheet, meaning the sheet's one Table.
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
  }
  private _ensureSheetState(): void {
    if (!this.sheetsStateRaw.has(this.sheetGid)) {
      this.sheetsStateRaw.set(
        this.sheetGid,
        emptyStateRaw.sheetState(this.sheetGid),
      );
    }
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
      return tableStateOf(this.tablesStateRaw, this.tableAddress.tableId);
    }
    const tableId = this.onlyTableId();
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
  hasOneTable(): boolean {
    return this.onlyTableId() !== undefined;
  }
  getRowState(rowIndex: number): RowStateRaw {
    return Val.assert(
      this.rowStates.get(rowIndex),
      `rowState for ${this.rowLabel(rowIndex)} on sheetGid ${this.sheetGid}`,
    );
  }
  // The live Table once fetched; before that, where the configs record it.
  tableOrigin(): TableOrigin {
    const properties = this.tableProperties;
    if (properties === undefined) return this.presumedOrigin;
    return originOf(properties);
  }
  get presumedOrigin(): TableOrigin {
    return new SpreadsheetSchema().presumedOrigin(this.sheetGid);
  }
  rowLabel(rowIndex: number): string {
    return `row ${this.tableOrigin().rowNumber(rowIndex)}`;
  }
  protected _integrateSheetProperties(sheet: SheetSnapshot): void {
    if (sheet.title) {
      this.sheetState.working.title = sheet.title;
    }
    if (sheet.rowCount !== undefined) {
      this.sheetState.working.rowCount = sheet.rowCount;
    }
    if (sheet.columnCount !== undefined) {
      this.sheetState.working.columnCount = sheet.columnCount;
    }
    if (sheet.tables !== undefined) {
      this._integrateTables(sheet.tables);
    }
    this._integrateQueuedSheetProperties();
  }
  // Every Table on the sheet is kept, so the one-Table-per-sheet refusal can count them.
  private _integrateTables(tables: TableSnapshot[]): void {
    const liveTableIds = tables.map(({ tableId }) => tableId);
    this.tableIds()
      .filter((tableId) => !liveTableIds.includes(tableId))
      .forEach((tableId) => this._removeAbsentTable(tableId));
    tables.forEach((table) => {
      const tableState = this._tableStateToIntegrate(
        table.tableId,
        tables.length === 1,
      );
      tableState.sheetGid = this.sheetGid;
      tableState.properties = {
        tableId: table.tableId,
        name: table.name,
        startRowIndex: table.startRowIndex,
        endRowIndex: table.endRowIndex,
        startColumnIndex: table.startColumnIndex,
        endColumnIndex: table.endColumnIndex,
        columnProperties: table.columnProperties,
        rowIndexesAreStale: tableState.properties?.rowIndexesAreStale ?? false,
      };
      integrateColumnProperties(tableState, table);
      this.tablesStateRaw.set(table.tableId, tableState);
    });
  }
  // A queued write must never vanish with its Table, so it stops the run instead.
  private _removeAbsentTable(tableId: string): void {
    const { writeQueue } = tableStateOf(this.tablesStateRaw, tableId);
    if (hasQueuedWrites(writeQueue)) {
      throw new Error(
        `Table ${tableId} is no longer on ${this.sheetLabel}, but it has queued writes; refetch before queuing writes to it.`,
      );
    }
    this.tablesStateRaw.delete(tableId);
  }
  // The sheet's one Table takes over what was queued through the sheet before it was known.
  private _tableStateToIntegrate(
    tableId: string,
    isSheetsOneTable: boolean,
  ): TableStateRaw {
    const existing = this.tablesStateRaw.get(tableId);
    if (existing !== undefined) {
      if (isSheetsOneTable) this._validateNoWritesQueuedThroughSheet();
      return existing;
    }
    if (!isSheetsOneTable) return emptyStateRaw.tableState(this.sheetGid);
    const adopted = this.sheetState.tableBeforeProperties;
    this.sheetState.tableBeforeProperties = emptyStateRaw.tableState(
      this.sheetGid,
    );
    return adopted;
  }
  // Once the sheet resolves to a Table it already knew, the flush no longer reads what was queued through the sheet.
  private _validateNoWritesQueuedThroughSheet(): void {
    if (hasQueuedWrites(this.sheetState.tableBeforeProperties.writeQueue)) {
      throw new Error(
        `Writes were queued through ${this.sheetLabel} while it had several Tables, and it now has one; refetch before queuing writes to it.`,
      );
    }
  }
  // Queued properties outlive a re-fetch until the flush sends them.
  private _integrateQueuedSheetProperties(): void {
    this.writeOperations.renameSheet.forEach(({ sheetId, title }) => {
      if (sheetId === this.sheetGid) this.sheetState.working.title = title;
    });
    this.writeOperations.renameTable.forEach(({ tableId, name }) => {
      this._updateWorkingTableName(tableId, name);
    });
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

function sheetGidOf(
  tables: TablesStateRaw,
  tableAddress: TableAddressRaw,
): number {
  if ("sheetGid" in tableAddress) return tableAddress.sheetGid;
  return tableStateOf(tables, tableAddress.tableId).sheetGid;
}

function tableStateOf(tables: TablesStateRaw, tableId: string): TableStateRaw {
  return Val.assert(tables.get(tableId), `Table state for tableId ${tableId}`);
}

function hasQueuedWrites({ table, rows }: TableWriteQueueRaw): boolean {
  return (
    rows.size > 0 ||
    table.sort !== undefined ||
    table.insertTableEndColumnCount > 0 ||
    table.fillColumns.length > 0 ||
    table.findReplaces.length > 0 ||
    table.columnTypes.size > 0
  );
}

export function originOf(properties: TablePropertiesRaw): TableOrigin {
  return new TableOrigin({
    headerRowIndex: properties.startRowIndex,
    startColIndex: properties.startColumnIndex,
  });
}

function integrateColumnProperties(
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
