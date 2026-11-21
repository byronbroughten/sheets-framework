import type { ConditionalFormatDeclaration } from "../00_Source/RawSource/ConditionalFormat.ts";
import type {
  EditLockDeclaration,
  EditWarningDeclaration,
} from "../00_Source/RawSource/EditProtection.ts";
import type { FindReplaceTerms } from "../00_Source/RawSource/RawSource.ts";
import type { Value } from "../01_SpreadsheetSchema/configReaders/valueSchemas.ts";
import { dimensionIds } from "../01_SpreadsheetSchema/dimensionIds.ts";
import { type HeadRole, headRows } from "../01_SpreadsheetSchema/headRows.ts";
import { TableRaw } from "../02_SpreadsheetRaw/TableRaw.ts";
import { TableCommonIdentified } from "./ClassBases/TableCommonIdentified.ts";
import { ColumnIdentified } from "./ColumnIdentified.ts";
import { HeadRowIdentified } from "./HeadRowIdentified.ts";
import { RowIdentified } from "./RowIdentified.ts";
import { SpreadsheetTableValidatorIdentified } from "./SpreadsheetIdentified/SpreadsheetTableValidatorIdentified.ts";
import { TableColumnResolverIdentified } from "./TableIdentified/TableColumnResolverIdentified.ts";

export class TableIdentified extends TableCommonIdentified {
  get raw(): TableRaw {
    return new TableRaw(this.tableIdentifiedProps);
  }
  get tableValidator(): SpreadsheetTableValidatorIdentified {
    return new SpreadsheetTableValidatorIdentified(this.tableIdentifiedProps);
  }
  get columnResolver(): TableColumnResolverIdentified {
    return new TableColumnResolverIdentified(this.tableIdentifiedProps);
  }
  get workingRowIndexes(): number[] {
    return this.raw.workingRowIndexes;
  }
  get rows(): RowIdentified[] {
    return this.raw.rows.map((row) => this.row(row.rowIndex));
  }
  get topRow(): RowIdentified {
    return this.row(0);
  }
  get rowCount(): number {
    return this.raw.rowCount;
  }
  // The invariant makes a zero row count permanently false, so ask this instead.
  get hasNoData(): boolean {
    if (this.raw.dataRowCountAfterFlush === 0) return true;
    return this._isTopRowTheOnlyRow && this.topRow.isBlank;
  }
  get workingRowIndexesWithData(): number[] {
    return this._withoutBlankRows(this.workingRowIndexes);
  }
  get rowIndexesFullWithData(): number[] {
    return this._withoutBlankRows(this.raw.rowIndexesFull);
  }
  get nonFormulaColumnIds(): string[] {
    return this._columnIdsOnSheet(this.schema.nonFormulaColumnIds);
  }
  // Feedback columns only report on a row, so a row holding nothing else is still blank.
  get blankTestColumnIds(): string[] {
    const feedbackColumnIds = this.feedbackColumnIds.get(this.schema.tableId);
    return this.nonFormulaColumnIds.filter(
      (columnId) => !feedbackColumnIds?.has(columnId),
    );
  }
  column(columnId: string): ColumnIdentified {
    return new ColumnIdentified({
      ...this.tableIdentifiedProps,
      columnId,
    });
  }
  gatherFetchDataPrepped(): void {
    // This is so that table dimensions and columnIndexes can be guaranteed
    // before their fetch requests are generated.
    this.fetchTargets.forEach((target) => {
      if (target.kind === "fullRow") {
        this.raw.rowCommon(target.row).gatherFetchFull();
      } else if (target.kind === "fullDataColumn") {
        this.column(target.column).raw.gatherFetchFull();
      } else if (target.kind === "singleCell") {
        const colIndex = this.column(target.column).colIndex;
        this.raw.rowCommon(target.row).cell(colIndex).gatherFetchRange();
      } else {
        const exhaustive: never = target;
        throw new Error(`Unknown fetch target: ${JSON.stringify(exhaustive)}`);
      }
    });
  }
  ensureColumnIdsAreFetched(): this {
    this.columnResolver.gatherDataPrerequisites();
    this.raw.ss.fetchAllGathered();
    this.tableValidator.validateUsedTables();
    return this;
  }
  addMissingColumnIds(): number {
    const colIndexes = this.raw.columnResolver.colIndexesWithoutColumnId;
    colIndexes.forEach((colIndex) => {
      this.raw
        .headRow("columnId")
        .updateValue(colIndex, dimensionIds.col(this.schema.idPrefix));
    });
    return colIndexes.length;
  }
  findReplace(terms: FindReplaceTerms): this {
    this.raw.findReplace(terms);
    return this;
  }
  addConditionalFormatRule(declaration: ConditionalFormatDeclaration): this {
    this.raw.addConditionalFormatRule(declaration);
    return this;
  }
  removeConditionalFormatRules(): this {
    this.raw.removeConditionalFormatRules();
    return this;
  }
  addEditWarning(declaration: EditWarningDeclaration = {}): this {
    this.raw.addEditWarning(declaration);
    return this;
  }
  addEditLock(declaration: EditLockDeclaration = {}): this {
    this.raw.addEditLock(declaration);
    return this;
  }
  removeEditProtections(): this {
    this.raw.removeEditProtections();
    return this;
  }
  anchoredA1(colIndex: number, rowIndex = 0): string {
    const origin = this.raw.tableOrigin();
    return this.schema.anchoredA1(
      origin.sheetColIndex(colIndex),
      origin.sheetRowIndex(rowIndex),
    );
  }
  row(rowIndex: number): RowIdentified {
    return new RowIdentified({
      ...this.tableIdentifiedProps,
      rowIndex,
    });
  }
  headRow<HR extends HeadRole>(headRole: HR): HeadRowIdentified<HR> {
    return new HeadRowIdentified({ ...this.tableIdentifiedProps, headRole });
  }
  headRowByIndex(rowIndex: number): HeadRowIdentified {
    return this.headRow(headRows.rolesAt(rowIndex)[0]);
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
    return this.row(this.raw.appendDataRow(this._formulaColIndexes).rowIndex);
  }
  private get _formulaColIndexes(): number[] {
    return this._columnIdsOnSheet(this.schema.formulaColumnIds).map(
      (columnId) => this.columnResolver.colIndexOf(columnId),
    );
  }
  // A configured column the sheet doesn't have holds nothing to read, clear, default or copy.
  private _columnIdsOnSheet(columnIds: string[]): string[] {
    return columnIds.filter((columnId) =>
      this.columnResolver.hasColumnId(columnId),
    );
  }
  private _isTopRowReusable(): boolean {
    if (!this._isTopRowTheOnlyRow) return false;
    if (!this.topRow.inWorking) {
      throw new Error(
        `Cannot append to ${this.raw.tableLabel}: its one data row was never fetched, so whether the append may reuse it is unknown. Prefetch its top data row first.`,
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
