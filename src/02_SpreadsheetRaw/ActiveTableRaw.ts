import type { TableColumnSnapshot } from "../00_Source/RawSource/RawSource";
import {
  type SheetColIndex,
  SheetIndex,
  type SheetRowIndex,
} from "../00_Source/RawSource/SheetIndex";
import { TableOrigin } from "../01_SpreadsheetSchema/TableOrigin";
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
  get origin(): TableOrigin {
    return new TableOrigin({
      headerRowIndex: this.startRowIndex,
      startColIndex: this.startColumnIndex,
    });
  }
  get startRowIndex(): SheetRowIndex {
    return this._knownTable().startRowIndex;
  }
  get startColumnIndex(): SheetColIndex {
    return this._knownTable().startColumnIndex;
  }
  get dataRowCount(): number {
    this.assertRowIndexesNotStale();
    const { startRowIndex, endRowIndex } = this._knownTable();
    return endRowIndex - startRowIndex - 1;
  }
  get columnCount(): number {
    const { startColumnIndex, endColumnIndex } = this._knownTable();
    return endColumnIndex - startColumnIndex;
  }
  get columnProperties(): TableColumnSnapshot[] {
    return this._knownTable().columnProperties;
  }
  get rowIndexesAreStale(): boolean {
    return this._knownTable().rowIndexesAreStale;
  }
  growDataRowCount(): void {
    this.assertRowIndexesNotStale();
    const knownTable = this._knownTable();
    knownTable.endRowIndex = SheetIndex.row(knownTable.endRowIndex + 1);
  }
  growColumnCount(): void {
    const knownTable = this._knownTable();
    knownTable.endColumnIndex = SheetIndex.col(knownTable.endColumnIndex + 1);
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
