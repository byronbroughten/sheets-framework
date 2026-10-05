import type {
  AddSheetOperation,
  BoundedGridRange,
  OpaqueRawRequest,
} from "../00_Source/RawSource/RawSource";
import { SpreadsheetSchema } from "../01_SpreadsheetSchema/SpreadsheetSchema";
import { validateFormulaString } from "./CellRaw";
import { SpreadsheetBaseRaw } from "./ClassBases/SpreadsheetBaseRaw";
import { emptyStateRaw } from "./ClassTypes/emptyStateRaw";
import type {
  AddedSheetCell,
  AddTableProps,
  FindReplaceProps,
} from "./ClassTypes/StateRaw";
import { SheetMetaRaw } from "./SheetMetaRaw";
import { SpreadsheetFetcherRaw } from "./SpreadsheetRaw/SpreadsheetFetcherRaw";
import { SpreadsheetFlusherRaw } from "./SpreadsheetRaw/SpreadsheetFlusherRaw";
import { TableRaw } from "./TableRaw";

/**
 * Spreadsheet-level Raw: GID+index fetch and the two Sheets chokepoints
 * (`fetchAllGathered` / `fetchSheetUsedGrid` via RawSource.fetchGrid,
 * `fetchAllSheetProperties` via RawSource.fetchSheetProperties,
 * `batchUpdateGSheets` via RawSource.flush), delegated to SpreadsheetRaw/.
 * A Table by tableId (or a sheet's one Table by GID), its rows and columns by
 * Table-relative index live on TableRaw / RowRaw / ColumnRaw here;
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
  get activeSheets(): TableRaw[] {
    return Array.from(this.activeSheetGids, (sheetGid) => this.sheet(sheetGid));
  }
  sheet(sheetGid: number): TableRaw {
    return new TableRaw({
      spreadsheetStateRaw: this.spreadsheetStateRaw,
      sheetGid: sheetGid,
    });
  }
  table(tableId: string): TableRaw {
    return new TableRaw({
      spreadsheetStateRaw: this.spreadsheetStateRaw,
      tableId,
    });
  }
  sheetMeta(sheetGid: number): SheetMetaRaw {
    return new SheetMetaRaw({
      spreadsheetStateRaw: this.spreadsheetStateRaw,
      sheetGid: sheetGid,
    });
  }
  sheets(...sheetGids: number[]): TableRaw[] {
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
  gatherAddSheetOperation(props: Omit<AddSheetOperation, "kind">): this {
    this.writeOperations.addSheet.push({ kind: "addSheet", ...props });
    return this;
  }
  // A created Table starts with a blank body row, the model growth copies from.
  gatherAddTableOperation({ tableId, ...props }: AddTableProps): this {
    const { startRowIndex, endRowIndex } = props.range;
    if (endRowIndex <= startRowIndex + 1) {
      throw new Error(
        `Add-Table refused: ${props.name}'s range holds only its header; a created Table needs a body row.`,
      );
    }
    this.writeOperations.addTable.push({
      kind: "addTable",
      tableId: tableId ?? randomTableId(),
      ...props,
    });
    return this;
  }
  // A seeded value on a tab this flush adds; an existing tab writes through CellRaw.
  gatherAddedSheetFillCellOperation(props: AddedSheetCell): this {
    this._validateAddSheetQueued(props.sheetId, "cell write");
    if ("formula" in props) validateFormulaString(props.formula);
    this.writeOperations.fillCell.push({ kind: "fillCell", ...props });
    return this;
  }
  // A checkbox on a tab this flush adds; an existing tab goes through CellRaw.
  gatherAddedSheetCheckboxValidationOperation(range: BoundedGridRange): this {
    this._validateAddSheetQueued(range.sheetId, "checkbox validation");
    this.writeOperations.addCheckboxValidation.push({
      kind: "addCheckboxValidation",
      range,
    });
    return this;
  }
  // Matches by content rather than by coordinate, so no local mirror is possible.
  findReplace({ scope, ...terms }: FindReplaceProps): this {
    this.writeOperations.findReplace.push({
      kind: "findReplace",
      terms,
      scope,
    });
    return this;
  }
  // The one bypass of the type layer; using it obliges filing an issue (docs/architecture/raw-request-opening.md).
  gatherRawOperation(request: OpaqueRawRequest): this {
    this.writeOperations.raw.push({ kind: "raw", request });
    return this;
  }
  // Abandons queued writes while local state still reflects them — terminal step only.
  discardQueuedChanges(): this {
    this.spreadsheetStateRaw.writeQueue = emptyStateRaw.spreadsheetWriteQueue();
    this.tablesStateRaw.forEach((state) => {
      state.writeQueue = emptyStateRaw.tableWriteQueue();
    });
    this.sheetsStateRaw.forEach((state) => {
      state.writeQueue = emptyStateRaw.sheetWriteQueue();
      state.tableBeforeProperties.writeQueue = emptyStateRaw.tableWriteQueue();
    });
    return this;
  }
  private _validateAddSheetQueued(sheetId: number, write: string): void {
    if (this.writeOperations.addSheet.some((op) => op.sheetId === sheetId)) {
      return;
    }
    throw new Error(
      `Added-sheet ${write} refused: no addSheet for GID ${sheetId} is queued in this flush.`,
    );
  }
}

function randomTableId(): string {
  const hexDigitCount = 10;
  let hexDigits = "";
  for (let i = 0; i < hexDigitCount; i++) {
    hexDigits += Math.floor(Math.random() * 16).toString(16);
  }
  return `tbl-${hexDigits}`;
}
