import type {
  GridFetchRange,
  SpreadsheetSnapshot,
} from "../../00_Source/RawSource/RawSource";
import { SpreadsheetSchema } from "../../01_SpreadsheetSchema/SpreadsheetSchema";
import { Val } from "../../utils/Val";
import { SpreadsheetBaseRaw } from "../ClassBases/SpreadsheetBaseRaw";
import { SpreadsheetRaw } from "../SpreadsheetRaw";
import {
  type MisplacedTable,
  type SheetIdentity,
  SpreadsheetTableValidatorRaw,
} from "./SpreadsheetTableValidatorRaw";

export class SpreadsheetFetcherRaw extends SpreadsheetBaseRaw {
  get ss(): SpreadsheetRaw {
    return new SpreadsheetRaw(this.spreadsheetRawProps);
  }
  get schema(): SpreadsheetSchema {
    return new SpreadsheetSchema();
  }
  get tableValidator(): SpreadsheetTableValidatorRaw {
    return new SpreadsheetTableValidatorRaw(this.spreadsheetRawProps);
  }
  ensureAllSheetPropertiesAreFetched(): void {
    if (!this.spreadsheetStateRaw.allSheetPropertiesAreFetched) {
      this._fetchAndIntegrateAllSheetProperties();
    }
  }
  ensureTimeZoneIsFetched(): string {
    if (this.spreadsheetStateRaw.timeZone === undefined) {
      this._fetchTimeZone();
    }
    return Val.assert(this.spreadsheetStateRaw.timeZone, "timeZone");
  }
  fetchAllSheetProperties(): { activeSheetGids: number[] } {
    this._fetchAndIntegrateAllSheetProperties();
    this.tableValidator.validateTablePlacement();
    return { activeSheetGids: this.ss.activeSheetGids };
  }
  fetchAllGathered(includeProgrammaticFacts = false): void {
    this._fetchGatheredConditionalFormatRules();
    this._fetchGatheredEditProtections();
    // An empty dataFilters list would fetch the whole spreadsheet's grid data.
    if (this.fetcherGridRanges.length === 0) return;
    const data = this._fetchByGridRanges(includeProgrammaticFacts);
    this._addDataToState(data);
    this._finalizeGatheredFetches();
    this.spreadsheetStateRaw.fetchQueue.gridRanges = [];
  }
  // One sheet by GID without Table-placement finalize, so a moved Table can wait for overlay.
  fetchSheetUsedGrid(sheetGid: number): void {
    const data = this._fetchByGridRanges(false, [{ sheetId: sheetGid }]);
    const sheets = data.sheets.filter((sheet) => sheet.sheetGid === sheetGid);
    if (sheets.length === 0) {
      throw new Error(`Sheet gid ${sheetGid} was missing from the Sheets get.`);
    }
    this._addDataToState({ ...data, sheets });
  }
  private _fetchTimeZone(): void {
    const timeZone = this.spreadsheetStateRaw.rawSource.fetchTimeZone();
    if (timeZone === null) {
      throw new Error(
        "The spreadsheet's properties.timeZone was absent from the Sheets get.",
      );
    }
    Logger.log(`Fetched the spreadsheet's time zone on its own: ${timeZone}.`);
    this.spreadsheetStateRaw.timeZone = timeZone;
  }
  private _fetchAndIntegrateAllSheetProperties(): void {
    const data = this.spreadsheetStateRaw.rawSource.fetchSheetProperties();
    this._addDataToState(data);
    this.spreadsheetStateRaw.allSheetPropertiesAreFetched = true;
  }
  // Backfills cells for every range fetched this cycle so a Sheets
  // response that omits empty cells (or whole blank rows) never leaves
  // them looking merely "not yet fetched" to callers.
  private _finalizeGatheredFetches(): void {
    const misplacedTables: MisplacedTable[] = [];
    const absentTables: SheetIdentity[] = [];
    this.spreadsheetStateRaw.sheets.forEach((state, sheetGid) => {
      // Above the early return, so a range that arrived incidentally is still judged.
      const placement = this.tableValidator.tablePlacement(sheetGid);
      if (placement.kind === "extra") {
        return;
      }
      if (placement.kind === "misplaced") {
        misplacedTables.push(placement);
        return;
      }
      const sheet = this.ss.sheet(sheetGid);
      const toFinalize = state.fetchQueue.toFinalize;
      sheet.finalizeFetchedCells();
      if (toFinalize.rows.size === 0 && toFinalize.columns.size === 0) {
        return;
      }
      if (state.working.knownTable === undefined) {
        absentTables.push({ sheetGid });
        return;
      }
      if (toFinalize.rows.has(this.schema.colIdRowIndex)) {
        state.working.hasFetchedColumnIds = true;
      }
      toFinalize.rows.forEach((rowIndex) => {
        sheet.rowCommon(rowIndex).ensureFullActiveDataCells();
      });
      toFinalize.columns.forEach((colIndex) => {
        sheet.column(colIndex).ensureFullActiveDataCells();
      });
      sheet.ensureFetchedActiveFacts();
      toFinalize.rows.clear();
      toFinalize.columns.clear();
    });
    if (absentTables.length > 0) {
      // The probe is built from the constants under test, so a moved Table looks absent.
      this.ensureAllSheetPropertiesAreFetched();
    }
    this.tableValidator.validateTablePlacement({
      misplacedTables,
      absentTables,
    });
  }
  // isFormula/numberFormatType (from rowData.values.userEnteredValue/
  // effectiveFormat) and column validation values/declared types (from
  // tables.columnProperties) are read only by ColumnConfigOperator's
  // programmatic value correction — every other caller only ever needs effectiveValue, so
  // those fields are left out of the default fetch to avoid fetching them
  // (and, for dataValidationRule, an unbounded list of validation values)
  // wastefully on every ordinary read.
  private _fetchByGridRanges(
    includeProgrammaticFacts: boolean,
    gridRanges: GridFetchRange[] = this.fetcherGridRanges,
  ): SpreadsheetSnapshot {
    return this.spreadsheetStateRaw.rawSource.fetchGrid(gridRanges, {
      includeProgrammaticFacts,
    });
  }
  private _fetchGatheredConditionalFormatRules(): void {
    const gatheringGids = this._gatheringGids("gatherConditionalFormats");
    if (gatheringGids.length === 0) return;
    this.spreadsheetStateRaw.rawSource
      .fetchConditionalFormatRules()
      .filter(({ sheetGid }) => gatheringGids.includes(sheetGid))
      .forEach(({ sheetGid, rules }) =>
        this.ss.sheet(sheetGid).integrateConditionalFormatRules(rules),
      );
  }
  private _fetchGatheredEditProtections(): void {
    const gatheringGids = this._gatheringGids("gatherEditProtections");
    if (gatheringGids.length === 0) return;
    this.spreadsheetStateRaw.rawSource
      .fetchEditProtections()
      .filter(({ sheetGid }) => gatheringGids.includes(sheetGid))
      .forEach(({ sheetGid, protections }) =>
        this.ss.sheet(sheetGid).integrateEditProtections(protections),
      );
  }
  private _gatheringGids(
    flag: "gatherConditionalFormats" | "gatherEditProtections",
  ): number[] {
    return Array.from(this.sheetsStateRaw.entries())
      .filter(([, state]) => state.fetchQueue[flag])
      .map(([sheetGid]) => sheetGid);
  }
  private _addDataToState(snapshot: SpreadsheetSnapshot): void {
    if (snapshot.timeZone !== null) {
      this.spreadsheetStateRaw.timeZone = snapshot.timeZone;
    }
    snapshot.sheets.forEach((sheetSnapshot) => {
      const sheet = this.ss.sheet(sheetSnapshot.sheetGid);
      sheet.integrateSheetState(sheetSnapshot);
    });
  }
}
