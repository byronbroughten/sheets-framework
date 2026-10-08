import type { SheetEdit } from "../00_Source/PlatformEvents/sheetEdit";
import { SpreadsheetSchema } from "../01_SpreadsheetSchema/SpreadsheetSchema";
import type { TableSchema } from "../01_SpreadsheetSchema/TableSchema";
import { SpreadsheetRaw } from "../02_SpreadsheetRaw/SpreadsheetRaw";
import { SpreadsheetBaseIdentified } from "./ClassBases/SpreadsheetBaseIdentified";
import { managedTableAddress } from "./ClassBases/TableBaseIdentified";
import { TableIdentified } from "./TableIdentified";
import type { GatherDataPrerequisitesProps } from "./TableIdentified/TableColumnResolverIdentified";

export class SpreadsheetIdentified extends SpreadsheetBaseIdentified {
  get schema(): SpreadsheetSchema {
    return new SpreadsheetSchema();
  }
  get raw(): SpreadsheetRaw {
    return new SpreadsheetRaw(this.spreadsheetRawProps);
  }
  tableOnSheet(sheetGid: number): TableIdentified {
    return new TableIdentified({
      ...this.spreadsheetIdentifiedProps,
      sheetGid,
    });
  }
  table(tableId: string): TableIdentified {
    return new TableIdentified({
      ...this.spreadsheetIdentifiedProps,
      tableId,
    });
  }
  // One fetch: the header zone brings every Table on the sheet with its column IDs.
  fetchTableWithActionCell(edit: SheetEdit): TableIdentified | undefined {
    if (!this.schema.mayHoldActionCell(edit)) return undefined;
    const tables = this.schema
      .tablesOnGid(edit.sheetGid)
      .map((tableSchema) => this.managedTable(tableSchema));
    tables.forEach((table) => table.columnResolver.gatherDataPrerequisites());
    this.raw.fetchAllGathered();
    return tables.find((table) =>
      table.raw.holdsActionCellAt(edit.rowIndexBase0, edit.colIndexBase0),
    );
  }
  managedTable(table: TableSchema): TableIdentified {
    return new TableIdentified({
      ...this.spreadsheetIdentifiedProps,
      ...managedTableAddress(table),
    });
  }
  get activeSheets(): TableIdentified[] {
    return this.raw.activeSheetGids.map((sheetGid) =>
      this.tableOnSheet(sheetGid),
    );
  }
  // Sheets first: building a sheet's handle hands its queue to its known Table.
  get tablesPreppedForFetch(): TableIdentified[] {
    return [...this._sheetsWaitingOnTable(), ...this._knownTables()].filter(
      (table) => table.isPreppedToFetch,
    );
  }
  // A rule or protection fetch is queued on the sheet, but it still needs its Table's column IDs.
  private _ensureGatheringSheetsHaveTables(): void {
    const knownGids = this._knownTables().map((table) => table.sheetGid);
    this.raw.activeSheetGids
      .filter((sheetGid) => this.raw.sheet(sheetGid).hasGatheredFetch)
      .filter((sheetGid) => !knownGids.includes(sheetGid))
      .forEach((sheetGid) => this.tableOnSheet(sheetGid));
  }
  fetchAllPrepped({
    includeProgrammaticFacts = false,
    ...props
  }: GatherDataPrerequisitesProps = {}): void {
    this._ensureGatheringSheetsHaveTables();
    const tablesPreppedForFetch = this.tablesPreppedForFetch;
    tablesPreppedForFetch.forEach((table) => {
      table.columnResolver.gatherDataPrerequisites(props);
    });
    this.raw.fetchAllGathered(includeProgrammaticFacts);
    tablesPreppedForFetch.forEach((table) => {
      table.gatherFetchDataPrepped();
    });
    this.raw.fetchAllGathered(includeProgrammaticFacts);
    tablesPreppedForFetch.forEach((table) => {
      table.clearFetchTargets();
    });
  }
  private _sheetsWaitingOnTable(): TableIdentified[] {
    return [...this.tableBeforePropertiesBySheet.keys()]
      .map((sheetGid) => this.tableOnSheet(sheetGid))
      .filter((table) => table.knownTableId === undefined);
  }
  private _knownTables(): TableIdentified[] {
    return [...this.tablesStateIdentified.keys()].map((tableId) =>
      this.table(tableId),
    );
  }
}
