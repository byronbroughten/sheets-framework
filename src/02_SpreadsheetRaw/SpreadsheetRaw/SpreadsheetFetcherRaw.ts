import type {
  GridFetchRange,
  SheetSnapshot,
  SpreadsheetSnapshot,
} from "../../00_Source/RawSource/RawSource";
import { SpreadsheetSchema } from "../../01_SpreadsheetSchema/SpreadsheetSchema";
import { Val } from "../../utils/Val";
import { SpreadsheetBaseRaw } from "../ClassBases/SpreadsheetBaseRaw";
import { emptyStateRaw } from "../ClassTypes/emptyStateRaw";
import { SpreadsheetRaw } from "../SpreadsheetRaw";
import {
  checkedTableIdsOf,
  type Misplacement,
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
    this._removeTablesAbsentFrom(sheets);
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
    this._removeTablesAbsentFrom(data.sheets);
    this._addDataToState(data);
    this.spreadsheetStateRaw.allSheetPropertiesAreFetched = true;
  }
  // Only a fetch covering the whole sheet lists every Table on it, so only it can tell one is gone.
  private _removeTablesAbsentFrom(sheets: SheetSnapshot[]): void {
    sheets.forEach((sheet) =>
      this.ss
        .tableOnSheet(sheet.sheetGid)
        .removeTablesAbsentFrom(sheet.tables ?? []),
    );
  }
  // Backfills cells for every range fetched this cycle so a Sheets
  // response that omits empty cells (or whole blank rows) never leaves
  // them looking merely "not yet fetched" to callers.
  private _finalizeGatheredFetches(): void {
    const misplacements: Misplacement[] = [];
    const headerOnlyTableIds: string[] = [];
    const finalizedSheetGids: number[] = [];
    const checkedTableIds: string[] = [];
    this.spreadsheetStateRaw.sheets.forEach((state, sheetGid) => {
      const placements = this.tableValidator.tablePlacements(sheetGid);
      state.tableBeforeProperties.fetchQueue = emptyStateRaw.tableFetchQueue();
      state.tablesBeforePropertiesById.forEach((tableState) => {
        tableState.fetchQueue = emptyStateRaw.tableFetchQueue();
      });
      state.fetchQueue.placementStripTableIds.clear();
      placements.forEach((placement) => {
        if (placement.kind === "misplaced") {
          misplacements.push(placement.misplacement);
        } else if (placement.kind === "header-only") {
          headerOnlyTableIds.push(placement.tableId);
        }
      });
      if (
        placements.every(
          ({ kind }) => kind === "none" || kind === "well-placed",
        )
      ) {
        finalizedSheetGids.push(sheetGid);
        checkedTableIds.push(...checkedTableIdsOf(placements));
      }
    });
    this.spreadsheetStateRaw.tables.forEach((state, tableId) => {
      if (!finalizedSheetGids.includes(state.sheetGid)) return;
      if (checkedTableIds.includes(tableId)) {
        state.working.placementIsChecked = true;
      }
      this.ss.table(tableId).finalizeFetches();
    });
    this.tableValidator.validateTablePlacement(
      misplacements,
      headerOnlyTableIds,
    );
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
      const table = this.ss.tableOnSheet(sheetSnapshot.sheetGid);
      table.integrateSheetState(sheetSnapshot);
    });
  }
}
