import { SpreadsheetSchema } from "../01_SpreadsheetSchema/SpreadsheetSchema";
import { SpreadsheetRaw } from "../02_SpreadsheetRaw/SpreadsheetRaw";
import { SpreadsheetBaseIdentified } from "./ClassBases/SpreadsheetBaseIdentified";
import { type ColumnIdentified } from "./ColumnIdentified";
import { SheetIdentified } from "./SheetIdentified";
import {
  type GatherDataPrerequisitesProps,
  SheetMetaIdentified,
} from "./SheetMetaIdentified";

export class SpreadsheetIdentified extends SpreadsheetBaseIdentified {
  get schema(): SpreadsheetSchema {
    return new SpreadsheetSchema();
  }
  get raw(): SpreadsheetRaw {
    return new SpreadsheetRaw(this.spreadsheetRawProps);
  }
  sheet(sheetGid: number): SheetIdentified {
    return new SheetIdentified({
      ...this.spreadsheetIdentifiedProps,
      sheetGid,
    });
  }
  sheetMeta(sheetGid: number): SheetMetaIdentified {
    return new SheetMetaIdentified({
      ...this.spreadsheetIdentifiedProps,
      sheetGid,
    });
  }
  column(sheetGid: number, columnId: string): ColumnIdentified {
    return this.sheet(sheetGid).column(columnId);
  }
  get activeSheets(): SheetIdentified[] {
    return this.raw.activeSheetGids.map((sheetGid) => this.sheet(sheetGid));
  }
  get sheetsPreppedForFetch(): SheetMetaIdentified[] {
    return Array.from(this.sheetsStateIdentified.keys())
      .map((sheetGid) => this.sheetMeta(sheetGid))
      .filter((sheet) => sheet.isPreppedToFetch);
  }
  fetchAllPrepped({
    includeProgrammaticFacts = false,
    ...props
  }: GatherDataPrerequisitesProps = {}): void {
    const sheetsPreppedForFetch = this.sheetsPreppedForFetch;
    sheetsPreppedForFetch.forEach((sheet) => {
      sheet._gatherDataPrerequisites(props);
    });
    this.raw.fetchAllGathered(includeProgrammaticFacts);
    sheetsPreppedForFetch.forEach((sheet) => {
      sheet.gatherFetchDataPrepped();
    });
    this.raw.fetchAllGathered(includeProgrammaticFacts);
    sheetsPreppedForFetch.forEach((sheet) => {
      sheet.clearFetchTargets();
    });
  }
}
