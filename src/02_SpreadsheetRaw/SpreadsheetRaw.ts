import type {
  AddSheetOperation,
  AddTableOperation,
  BoundedGridRange,
  OpaqueRawRequest,
} from "../00_Source/RawSource/RawSource";
import { SpreadsheetSchema } from "../01_SpreadsheetSchema/SpreadsheetSchema";
import { validateFormulaString } from "./CellRaw";
import { SpreadsheetBaseRaw } from "./ClassBases/SpreadsheetBaseRaw";
import { emptyStateRaw } from "./ClassTypes/emptyStateRaw";
import type { AddedSheetCell, FindReplaceProps } from "./ClassTypes/StateRaw";
import { SheetMetaRaw } from "./SheetMetaRaw";
import { SheetRaw } from "./SheetRaw";
import { SpreadsheetFetcherRaw } from "./SpreadsheetRaw/SpreadsheetFetcherRaw";
import { SpreadsheetFlusherRaw } from "./SpreadsheetRaw/SpreadsheetFlusherRaw";

/**
 * Spreadsheet-level Raw: GID+index fetch and the two Sheets chokepoints
 * (`fetchAllGathered` / `fetchSheetUsedGrid` via RawSource.fetchGrid,
 * `fetchAllSheetProperties` via RawSource.fetchSheetProperties,
 * `batchUpdateGSheets` via RawSource.flush), delegated to SpreadsheetRaw/.
 * Sheet/row/column by index live on SheetRaw / RowRaw / ColumnRaw here;
 * by-name and columnId resolution are Identified/Named. Schema classes that
 * resolve columns live in Schema/ because they sit below both consumer tiers.
 * docs/architecture/round-trips.md, schema-classes.md, class-chains.md
 */
export class SpreadsheetRaw extends SpreadsheetBaseRaw {
  static init(): SpreadsheetRaw {
    return new SpreadsheetRaw(SpreadsheetBaseRaw.initSpreadsheetRawProps());
  }
  get schema(): SpreadsheetSchema {
    return new SpreadsheetSchema();
  }
  private get fetcher(): SpreadsheetFetcherRaw {
    return new SpreadsheetFetcherRaw(this.spreadsheetRawProps);
  }
  private get flusher(): SpreadsheetFlusherRaw {
    return new SpreadsheetFlusherRaw(this.spreadsheetRawProps);
  }
  get timeZone(): string {
    return this.fetcher.ensureTimeZoneIsFetched();
  }
  gidIsActive(sheetGid: number): boolean {
    return this.activeSheetGids.includes(sheetGid);
  }
  get activeSheetGids(): number[] {
    return Array.from(this.spreadsheetStateRaw.sheets.keys());
  }
  get activeSheets(): SheetRaw[] {
    return Array.from(this.activeSheetGids, (sheetGid) => this.sheet(sheetGid));
  }
  sheet(sheetGid: number): SheetRaw {
    return new SheetRaw({
      spreadsheetStateRaw: this.spreadsheetStateRaw,
      sheetGid: sheetGid,
    });
  }
  sheetMeta(sheetGid: number): SheetMetaRaw {
    return new SheetMetaRaw({
      spreadsheetStateRaw: this.spreadsheetStateRaw,
      sheetGid: sheetGid,
    });
  }
  sheets(...sheetGids: number[]): SheetRaw[] {
    return sheetGids.map((sheetGid) => this.sheet(sheetGid));
  }
  ensureAllSheetPropertiesAreFetched(): void {
    this.fetcher.ensureAllSheetPropertiesAreFetched();
  }
  fetchAllSheetProperties(): { activeSheetGids: number[] } {
    return this.fetcher.fetchAllSheetProperties();
  }
  fetchAllGathered(includeProgrammaticFacts = false): void {
    this.fetcher.fetchAllGathered(includeProgrammaticFacts);
  }
  fetchSheetUsedGrid(sheetGid: number): void {
    this.fetcher.fetchSheetUsedGrid(sheetGid);
  }
  batchUpdateGSheets(): void {
    this.flusher.flush();
  }
  // Queued on the spreadsheet: a tab that does not exist yet has no sheet state to hold it.
  gatherAddSheetRequest(props: Omit<AddSheetOperation, "kind">): this {
    this.updateRequests.addSheet.push({ kind: "addSheet", ...props });
    return this;
  }
  gatherAddTableRequest(props: Omit<AddTableOperation, "kind">): this {
    this.updateRequests.addTable.push({ kind: "addTable", ...props });
    return this;
  }
  // A seeded value on a tab this flush adds; an existing tab writes through CellRaw.
  gatherAddedSheetCellRequest(props: AddedSheetCell): this {
    this._validateAddSheetQueued(props.sheetId, "cell write");
    if ("formula" in props) validateFormulaString(props.formula);
    this.updateRequests.update.push({ kind: "updateCell", ...props });
    return this;
  }
  // A checkbox on a tab this flush adds; an existing tab goes through CellRaw.
  gatherAddedSheetCheckboxValidationRequest(range: BoundedGridRange): this {
    this._validateAddSheetQueued(range.sheetId, "checkbox validation");
    this.updateRequests.addCheckboxValidation.push({
      kind: "addCheckboxValidation",
      range,
    });
    return this;
  }
  // Matches by content rather than by coordinate, so no local mirror is possible.
  findReplace({ scope, ...terms }: FindReplaceProps): this {
    this.updateRequests.findReplace.push({
      kind: "findReplace",
      terms,
      scope,
    });
    return this;
  }
  // The one bypass of the type layer; using it obliges filing an issue (docs/architecture/raw-request-opening.md).
  gatherRawRequest(request: OpaqueRawRequest): this {
    this.updateRequests.raw.push({ kind: "raw", request });
    return this;
  }
  // Abandons queued writes while local state still reflects them — terminal step only.
  discardQueuedChanges(): this {
    this.spreadsheetStateRaw.writeQueue = emptyStateRaw.spreadsheetWriteQueue();
    this.sheetsStateRaw.forEach((state) => {
      state.writeQueue = emptyStateRaw.sheetWriteQueue();
    });
    return this;
  }
  private _validateAddSheetQueued(sheetId: number, write: string): void {
    if (this.updateRequests.addSheet.some((op) => op.sheetId === sheetId)) {
      return;
    }
    throw new Error(
      `Added-sheet ${write} refused: no addSheet for GID ${sheetId} is queued in this flush.`,
    );
  }
}
