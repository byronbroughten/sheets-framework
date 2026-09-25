import type { TableColumnSnapshot } from "../00_Source/RawSource/RawSource";
import { Val } from "../utils/Val";
import type { SheetRawProps } from "./ClassBases/SheetBaseRaw";
import type {
  KnownTableRaw,
  SheetStateRaw,
  StateRaw,
} from "./ClassTypes/StateRaw";

export class ActiveTableRaw {
  readonly sheetGid: number;
  private readonly spreadsheetStateRaw: StateRaw;
  constructor({ sheetGid, spreadsheetStateRaw }: SheetRawProps) {
    this.sheetGid = sheetGid;
    this.spreadsheetStateRaw = spreadsheetStateRaw;
  }
  get tableId(): string {
    return this._knownTable().tableId;
  }
  get name(): string {
    return this._knownTable().name;
  }
  get startRowIndex(): number {
    return this._knownTable().startRowIndex;
  }
  get endRowIndex(): number {
    this.assertRowIndexesNotStale();
    return this._knownTable().endRowIndex;
  }
  set endRowIndex(endRowIndex: number) {
    this.assertRowIndexesNotStale();
    this._knownTable().endRowIndex = endRowIndex;
  }
  get startColumnIndex(): number {
    return this._knownTable().startColumnIndex;
  }
  get endColumnIndex(): number {
    return this._knownTable().endColumnIndex;
  }
  set endColumnIndex(endColumnIndex: number) {
    this._knownTable().endColumnIndex = endColumnIndex;
  }
  get columnProperties(): TableColumnSnapshot[] {
    return this._knownTable().columnProperties;
  }
  get rowIndexesAreStale(): boolean {
    return this._knownTable().rowIndexesAreStale;
  }
  growEndRowIndex(): void {
    this.endRowIndex++;
  }
  growEndColumnIndex(): void {
    this.endColumnIndex++;
  }
  markRowIndexesStale(): void {
    this._knownTable().rowIndexesAreStale = true;
  }
  clearRowIndexStale(): void {
    this._knownTable().rowIndexesAreStale = false;
  }
  // The sent list is now the Table's, and what was fetched no longer is.
  markColumnPropertiesStale(): void {
    this._knownTable().columnProperties = [];
  }
  assertKnown(): void {
    this._knownTable();
  }
  assertRowIndexesNotStale(): void {
    if (!this._knownTable().rowIndexesAreStale) return;
    throw new Error(`Row indexes are stale for sheetGid ${this.sheetGid}.`);
  }
  ensureColIndexIsStale(colIndex: number): void {
    const knownTable = this._knownTable();
    knownTable.firstStaleColIndex = Math.min(
      knownTable.firstStaleColIndex ?? Infinity,
      colIndex,
    );
  }
  validateColIndexNotStale(colIndex: number): void {
    const { firstStaleColIndex } = this._knownTable();
    if (firstStaleColIndex !== undefined && colIndex >= firstStaleColIndex) {
      throw new Error(
        `Column index ${colIndex} is stale. First stale column index is ${firstStaleColIndex}.`,
      );
    }
  }
  private get sheetState(): SheetStateRaw {
    return Val.assert(
      this.spreadsheetStateRaw.sheets.get(this.sheetGid),
      `sheetState for sheetGid ${this.sheetGid}`,
    );
  }
  private _knownTable(): KnownTableRaw {
    const knownTable = this.sheetState.working.knownTable;
    if (knownTable === undefined) {
      throw new Error(
        `Active table is null for sheetGid ${this.sheetGid}. Ensure that the sheet properties have been fetched.`,
      );
    }
    return knownTable;
  }
}
