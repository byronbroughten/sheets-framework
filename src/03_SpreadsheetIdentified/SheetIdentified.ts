import type {
  ConditionalFormatDeclaration,
  ConditionalFormatRule,
} from "../00_Source/RawSource/ConditionalFormat";
import type {
  EditLockDeclaration,
  EditProtection,
  EditWarningDeclaration,
  WholeSheetEditLockDeclaration,
  WholeSheetEditWarningDeclaration,
} from "../00_Source/RawSource/EditProtection";
import type { Value } from "../01_SpreadsheetSchema/valueSchemas";
import type { FindReplaceTerms } from "../02_SpreadsheetRaw/ClassTypes/StateRaw";
import { SheetRaw } from "../02_SpreadsheetRaw/SheetRaw";
import { SheetCommonIdentified } from "./ClassBases/SheetCommonIdentified";
import { ColumnIdentified } from "./ColumnIdentified";
import { RowIdentified } from "./RowIdentified";
import { SheetMetaIdentified } from "./SheetMetaIdentified";

export class SheetIdentified extends SheetCommonIdentified {
  get meta(): SheetMetaIdentified {
    return new SheetMetaIdentified(this.sheetIdentifiedProps);
  }
  get raw(): SheetRaw {
    return new SheetRaw(this.sheetIdentifiedProps);
  }
  get rowIndexesActive(): number[] {
    return this.raw.rowIndexesActive;
  }
  get rows(): RowIdentified[] {
    return this.raw.rows.map((row) => this.row(row.rowIndex));
  }
  get topRow(): RowIdentified {
    return this.row(this.schema.topDataRowIdx);
  }
  get rowCount(): number {
    return this.raw.rowCount;
  }
  // The invariant makes a zero row count permanently false, so ask this instead.
  get hasNoData(): boolean {
    if (this.raw.dataRowCountAfterFlush === 0) return true;
    return this._isTopRowTheOnlyRow && this.topRow.isBlank;
  }
  get rowIndexesActiveWithData(): number[] {
    return this._withoutBlankRows(this.rowIndexesActive);
  }
  get rowIndexesFullWithData(): number[] {
    return this._withoutBlankRows(this.raw.rowIndexesFull);
  }
  // A configured column the sheet doesn't have holds nothing to read, clear, or default.
  get nonFormulaColumnIds(): string[] {
    return this.schema.nonFormulaColumnIds.filter((columnId) =>
      this.meta.isActiveColumnId(columnId),
    );
  }
  column(columnId: string): ColumnIdentified {
    return new ColumnIdentified({
      ...this.sheetIdentifiedProps,
      columnId,
    });
  }
  findReplace(terms: FindReplaceTerms): this {
    this.raw.findReplace(terms);
    return this;
  }
  prepFetchConditionalFormatRules(): this {
    this.sheetState.fetchQueue.gatherConditionalFormats = true;
    return this;
  }
  conditionalFormatRules(): ConditionalFormatRule[] {
    return this.raw.conditionalFormatRules();
  }
  addConditionalFormatRule(declaration: ConditionalFormatDeclaration): this {
    this.raw.addConditionalFormatRule(declaration);
    return this;
  }
  removeConditionalFormatRules(): this {
    this.raw.removeConditionalFormatRules();
    return this;
  }
  removeConditionalFormatRule(rule: ConditionalFormatRule): this {
    this.raw.removeConditionalFormatRule(rule);
    return this;
  }
  prepFetchEditProtections(): this {
    this.sheetState.fetchQueue.gatherEditProtections = true;
    return this;
  }
  editProtections(): EditProtection[] {
    return this.raw.editProtections();
  }
  addEditWarning(declaration: EditWarningDeclaration = {}): this {
    this.raw.addEditWarning(declaration);
    return this;
  }
  addEditLock(declaration: EditLockDeclaration = {}): this {
    this.raw.addEditLock(declaration);
    return this;
  }
  addEditWarningWholeSheet(
    declaration: WholeSheetEditWarningDeclaration = {},
  ): this {
    this.raw.addEditWarningWholeSheet(declaration);
    return this;
  }
  addEditLockWholeSheet(declaration: WholeSheetEditLockDeclaration = {}): this {
    this.raw.addEditLockWholeSheet(declaration);
    return this;
  }
  removeEditProtections(): this {
    this.raw.removeEditProtections();
    return this;
  }
  removeEditProtection(protection: EditProtection): this {
    this.raw.removeEditProtection(protection);
    return this;
  }
  removeEditProtectionByDescription(description: string): this {
    this.raw.removeEditProtectionByDescription(description);
    return this;
  }
  removeEditProtectionById(protectionId: number): this {
    this.raw.removeEditProtectionById(protectionId);
    return this;
  }
  anchoredA1(colIndex: number): string {
    return this.schema.anchoredA1(colIndex, this.schema.topDataRowIdx);
  }
  row(rowIndex: number): RowIdentified {
    return new RowIdentified({
      ...this.sheetIdentifiedProps,
      rowIndex,
    });
  }
  // The top data row survives, so which row a wipe leaves is predictable.
  DELETE_ALL_DATA_ROWS(): void {
    const [topRowIndex, ...rowIndexesBelow] = this.raw.rowIndexesFull;
    if (topRowIndex === undefined) return;
    rowIndexesBelow.forEach((rowIndex) => this.row(rowIndex).delete());
    this.row(topRowIndex).clearValues();
  }
  appendRowDefault(): RowIdentified {
    const row = this._rowToFillWithDefaults();
    this._defaultDataValues().forEach((value, columnId) => {
      row.updateValue(columnId, value);
    });
    row.reserve();
    return row;
  }
  // A queued delete makes the survivor's index unknowable until the flush.
  private get _isTopRowTheOnlyRow(): boolean {
    return (
      this.raw.dataRowCountAfterFlush === 1 && !this.topRow.isQueuedForDelete
    );
  }
  private _withoutBlankRows(rowIndexes: number[]): number[] {
    return rowIndexes.filter((rowIndex) => !this.row(rowIndex).isBlank);
  }
  // Appending past a blank row would leave it stranded above the data forever.
  private _rowToFillWithDefaults(): RowIdentified {
    if (this._isTopRowReusable()) return this.topRow;
    return this.row(this.raw.appendDataRow().rowIndex);
  }
  private _isTopRowReusable(): boolean {
    if (!this._isTopRowTheOnlyRow) return false;
    if (!this.topRow.isActive) {
      throw new Error(
        `Cannot append to sheetGid ${this.sheetGid}: its one data row was never fetched, so whether the append may reuse it is unknown. Prefetch that row first.`,
      );
    }
    return this.topRow.isReusable;
  }
  private _defaultDataValues(): Map<string, Value> {
    return this.nonFormulaColumnIds.reduce((acc, columnId) => {
      acc.set(
        columnId,
        this.schema.columnById(columnId).makeDefaultDataValue(),
      );
      return acc;
    }, new Map<string, Value>());
  }
}
