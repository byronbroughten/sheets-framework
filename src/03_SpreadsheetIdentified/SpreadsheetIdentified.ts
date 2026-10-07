import { SpreadsheetSchema } from "../01_SpreadsheetSchema/SpreadsheetSchema";
import { SpreadsheetRaw } from "../02_SpreadsheetRaw/SpreadsheetRaw";
import { SpreadsheetBaseIdentified } from "./ClassBases/SpreadsheetBaseIdentified";
import { SheetMetaIdentified } from "./SheetMetaIdentified";
import { TableIdentified } from "./TableIdentified";
import type { GatherDataPrerequisitesProps } from "./TableIdentified/TableColumnResolverIdentified";

export class SpreadsheetIdentified extends SpreadsheetBaseIdentified {
  get schema(): SpreadsheetSchema {
    return new SpreadsheetSchema();
  }
  get raw(): SpreadsheetRaw {
    return new SpreadsheetRaw(this.spreadsheetRawProps);
  }
  sheetMeta(sheetGid: number): SheetMetaIdentified {
    return new SheetMetaIdentified({
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
    return this.raw.activeSheetGids.map(
      (sheetGid) => this.sheetMeta(sheetGid).primary,
    );
  }
  // Sheets first: building a sheet's handle hands its queue to its known Table.
  get tablesPreppedForFetch(): SheetMetaIdentified[] {
    return [...this._sheetsWaitingOnTable(), ...this._knownTables()].filter(
      (table) => table.isPreppedToFetch,
    );
  }
  fetchAllPrepped({
    includeProgrammaticFacts = false,
    ...props
  }: GatherDataPrerequisitesProps = {}): void {
    const tablesPreppedForFetch = this.tablesPreppedForFetch;
    tablesPreppedForFetch.forEach((table) => {
      table.primary.columnResolver.gatherDataPrerequisites(props);
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
  private _sheetsWaitingOnTable(): SheetMetaIdentified[] {
    return [...this.tableBeforePropertiesBySheet.keys()]
      .map((sheetGid) => this.sheetMeta(sheetGid))
      .filter((sheet) => sheet.knownTableId === undefined);
  }
  private _knownTables(): SheetMetaIdentified[] {
    return [...this.tablesStateIdentified.keys()].map((tableId) =>
      this._tableMeta(tableId),
    );
  }
  private _tableMeta(tableId: string): SheetMetaIdentified {
    return new SheetMetaIdentified({
      ...this.spreadsheetIdentifiedProps,
      tableId,
    });
  }
}
