import type {
  SheetSnapshot,
  TableSnapshot,
} from "../../00_Source/RawSource/RawSource";
import { Obj } from "../../utils/Obj";
import { Val } from "../../utils/Val";
import { emptyStateRaw } from "../ClassTypes/emptyStateRaw";
import type {
  ColumnStateRaw,
  RowStateRaw,
  SheetStateRaw,
} from "../ClassTypes/StateRaw";
import {
  SpreadsheetBaseRaw,
  type SpreadsheetRawProps,
} from "./SpreadsheetBaseRaw";

export interface SheetRawProps extends SpreadsheetRawProps {
  sheetGid: number;
}

export class SheetBaseRaw extends SpreadsheetBaseRaw {
  readonly sheetGid: number;
  constructor({ sheetGid, ...rest }: SheetRawProps) {
    super(rest);
    this.sheetGid = sheetGid;
    this._ensureSheetState();
  }
  private _ensureSheetState(): void {
    if (!this.spreadsheetStateRaw.sheets.has(this.sheetGid)) {
      this.spreadsheetStateRaw.sheets.set(
        this.sheetGid,
        emptyStateRaw.sheetState(),
      );
    }
  }
  protected _initSheetState(sheet: SheetSnapshot): void {
    this._integrateSheetProperties(sheet);
    this._integrateQueuedSheetProperties();
  }
  private _integrateSheetProperties(sheet: SheetSnapshot): void {
    if (sheet.title) {
      this.sheetState.working.title = sheet.title;
    }
    const tables = sheet.tables;
    if (!tables) {
      return;
    }
    this.sheetState.working.tables = tables.map((table) => ({
      tableId: table.tableId,
      name: table.name,
    }));
    if (tables.length > 1) {
      this.sheetState.working.hasExtraTables = true;
      this.sheetState.working.knownTable = undefined;
      this._clearColumnPropertyFields();
      return;
    }
    this.sheetState.working.hasExtraTables = false;
    if (tables.length === 0) {
      this.sheetState.working.knownTable = undefined;
      return;
    }
    const table = Val.assert(tables[0], "table");
    const range = Obj.validatePick(
      table,
      "number",
      "startRowIndex",
      "endRowIndex",
      "startColumnIndex",
      "endColumnIndex",
    );
    const previous = this.sheetState.working.knownTable;
    this.sheetState.working.knownTable = {
      tableId: table.tableId,
      name: table.name,
      ...range,
      columnProperties: table.columnProperties,
      rowIndexesAreStale: previous?.rowIndexesAreStale ?? false,
      firstStaleColIndex: previous?.firstStaleColIndex,
    };
    this._parseColumnProperties(table, range.startColumnIndex);
  }
  // Queued properties outlive a re-fetch until the flush sends them.
  private _integrateQueuedSheetProperties(): void {
    const working = this.sheetState.working;
    this.updateRequests.updateSheetTitle.forEach(({ sheetId, title }) => {
      if (sheetId === this.sheetGid) working.title = title;
    });
    this.updateRequests.updateTableName.forEach(({ tableId, name }) => {
      this._updateWorkingTableName(tableId, name);
    });
    const knownTable = working.knownTable;
    if (knownTable === undefined) return;
    this.updateRequests.updateTableColumnType.forEach(
      ({ tableId, columnIndex, columnType }) => {
        if (tableId !== knownTable.tableId) return;
        this._ensureColumnState(
          knownTable.startColumnIndex + columnIndex,
        ).columnType = columnType;
      },
    );
  }
  protected _updateWorkingTableName(tableId: string, name: string): void {
    const working = this.sheetState.working;
    working.tables = working.tables.map((table) =>
      table.tableId === tableId ? { ...table, name } : table,
    );
    if (working.knownTable?.tableId === tableId) {
      working.knownTable.name = name;
    }
  }
  private _clearColumnPropertyFields(): void {
    this.sheetState.working.columnStates.forEach((columnState) => {
      delete columnState.validationValues;
      delete columnState.validationConditionType;
      delete columnState.columnType;
    });
  }
  private _parseColumnProperties(
    table: TableSnapshot,
    startColumnIndex: number,
  ): void {
    this._clearColumnPropertyFields();
    table.columnProperties.forEach((colProps) => {
      // The API states columnIndex table-relative.
      const colIndex = startColumnIndex + colProps.columnIndex;
      const columnState = this._ensureColumnState(colIndex);
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
  }
  protected _ensureColumnState(colIndex: number): ColumnStateRaw {
    const existing = this.sheetState.working.columnStates.get(colIndex);
    if (existing !== undefined) return existing;
    const created: ColumnStateRaw = {};
    this.sheetState.working.columnStates.set(colIndex, created);
    return created;
  }
  protected get sheetState(): SheetStateRaw {
    return Val.assert(
      this.spreadsheetStateRaw.sheets.get(this.sheetGid),
      `sheetState for sheetGid ${this.sheetGid}`,
    );
  }
  getRowState(rowIndex: number): RowStateRaw {
    return Val.assert(
      this.sheetState.working.rowStates.get(rowIndex),
      `rowState for row ${rowIndex} on sheetGid ${this.sheetGid}`,
    );
  }
  get columnStates(): SheetStateRaw["working"]["columnStates"] {
    return this.sheetState.working.columnStates;
  }
  get sheetLabel(): string {
    return `"${this.sheetState.working.title ?? "(untitled)"}" (gid ${this.sheetGid})`;
  }
  get rowStates(): SheetStateRaw["working"]["rowStates"] {
    return this.sheetState.working.rowStates;
  }
  get sheetRawProps(): SheetRawProps {
    return {
      sheetGid: this.sheetGid,
      ...this.spreadsheetRawProps,
    };
  }
}
