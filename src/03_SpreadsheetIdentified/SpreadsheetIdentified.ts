import type { SheetEdit } from "../00_Source/PlatformEvents/sheetEdit";
import { SpreadsheetSchema } from "../01_SpreadsheetSchema/configReaders/SpreadsheetSchema";
import { TableSchema } from "../01_SpreadsheetSchema/configReaders/TableSchema";
import { SpreadsheetRaw } from "../02_SpreadsheetRaw/SpreadsheetRaw";
import { SpreadsheetBaseIdentified } from "./ClassBases/SpreadsheetBaseIdentified";
import { SpreadsheetTableValidatorIdentified } from "./SpreadsheetIdentified/SpreadsheetTableValidatorIdentified";
import { TableIdentified } from "./TableIdentified";
import type { GatherDataPrerequisitesProps } from "./TableIdentified/TableColumnResolverIdentified";

export class SpreadsheetIdentified extends SpreadsheetBaseIdentified {
  get schema(): SpreadsheetSchema {
    return new SpreadsheetSchema();
  }
  get raw(): SpreadsheetRaw {
    return new SpreadsheetRaw(this.spreadsheetRawProps);
  }
  get tableValidator(): SpreadsheetTableValidatorIdentified {
    return new SpreadsheetTableValidatorIdentified(
      this.spreadsheetIdentifiedProps,
    );
  }
  // The one managed Table the configs record on that sheet.
  tableOnSheet(sheetGid: number): TableIdentified {
    const table = this.schema.loneTableOnGid(sheetGid);
    if (table === undefined) {
      throw new Error(
        `The configs record no Table or several on sheetGid ${sheetGid}; reach a shared sheet's Tables by Table ID.`,
      );
    }
    return this.managedTable(table);
  }
  table(tableId: string): TableIdentified {
    return this.managedTable(TableSchema.fromTableId(tableId));
  }
  // One fetch: the header zone brings every Table on the sheet with its column IDs, and only the ticked one is judged.
  fetchTableWithActionCell(edit: SheetEdit): TableIdentified | undefined {
    if (!this.schema.mayHoldActionCell(edit)) return undefined;
    this.raw.sheet(edit.sheetGid).gatherFetchHeaderZone();
    this.fetchAllGathered();
    const table = this.schema
      .tablesOnGid(edit.sheetGid)
      .map((tableSchema) => this.managedTable(tableSchema))
      .find((managedTable) =>
        managedTable.raw.holdsActionCellAt(
          edit.rowIndexBase0,
          edit.colIndexBase0,
        ),
      );
    if (table === undefined) return undefined;
    this.tableValidator.validateTables([table.schema]);
    table.raw.integrateHeaderZone();
    return table;
  }
  // Judges the Tables the fetch used once it has finalized them all.
  fetchAllGathered(includeProgrammaticFacts = false): void {
    this.raw.fetchAllGathered(includeProgrammaticFacts);
    this.tableValidator.validateUsedTables();
  }
  batchUpdateGSheets(): void {
    this.tableValidator.validateTablesForColumnInserts();
    this.raw.batchUpdateGSheets();
  }
  managedTable(table: TableSchema): TableIdentified {
    return new TableIdentified({
      ...this.spreadsheetIdentifiedProps,
      tableId: table.tableId,
      sheetGid: table.sheetGid,
    });
  }
  get activeSheets(): TableIdentified[] {
    return this.raw.activeSheetGids.map((sheetGid) =>
      this.tableOnSheet(sheetGid),
    );
  }
  get tablesPreppedForFetch(): TableIdentified[] {
    return this._knownTables().filter((table) => table.isPreppedToFetch);
  }
  // A rule or protection fetch is queued on the sheet, but it still needs its Table's column IDs.
  private _ensureGatheringSheetsHaveTables(): void {
    const knownGids = this._knownTables().map((table) => table.sheetGid);
    this.raw.activeSheetGids
      .filter((sheetGid) => this.raw.sheet(sheetGid).hasGatheredFetch)
      .filter((sheetGid) => !knownGids.includes(sheetGid))
      .flatMap((sheetGid) => this.schema.tablesOnGid(sheetGid))
      .forEach((table) => this.managedTable(table));
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
    this.fetchAllGathered(includeProgrammaticFacts);
    tablesPreppedForFetch.forEach((table) => {
      table.gatherFetchDataPrepped();
    });
    this.fetchAllGathered(includeProgrammaticFacts);
    tablesPreppedForFetch.forEach((table) => {
      table.clearFetchTargets();
    });
  }
  private _knownTables(): TableIdentified[] {
    return [...this.tablesStateIdentified.keys()].map((tableId) =>
      this.table(tableId),
    );
  }
}
