import { SpreadsheetSchema } from "../01_SpreadsheetSchema/SpreadsheetSchema";
import { SpreadsheetRaw } from "../02_SpreadsheetRaw/SpreadsheetRaw";
import { SpreadsheetBaseIdentified } from "./ClassBases/SpreadsheetBaseIdentified";
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
