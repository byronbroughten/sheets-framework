import type { CellValueName } from "../00_Source/CellValues/cellValues";
import type {
  ConditionalFormatDeclaration,
  ConditionalFormatRule,
} from "../00_Source/RawSource/ConditionalFormat";
import {
  type EditLockDeclaration,
  type EditProtection,
  type EditWarningDeclaration,
  type ProtectionGridRange,
  type WholeSheetEditLockDeclaration,
  type WholeSheetEditWarningDeclaration,
} from "../00_Source/RawSource/EditProtection";
import type {
  BoundedGridRange,
  GridRangeProps,
  SheetSnapshot,
  TableColumnPropertiesUpdate,
  TableColumnSnapshot,
} from "../00_Source/RawSource/RawSource";
import type { Value } from "../01_SpreadsheetSchema/valueSchemas";
import { Arr } from "../utils/Arr";
import { Val } from "../utils/Val";
import { assertValueAndFormulaExclusive } from "./CellRaw";
import type { RowCommonRaw } from "./ClassBases/RowCommonRaw";
import { SheetCommonRaw } from "./ClassBases/SheetCommonRaw";
import {
  type ColumnFill,
  type FindReplaceTerms,
  type SortParameters,
  type TableIdentityRaw,
  type UpdateTableColumnTypeOperation,
} from "./ClassTypes/StateRaw";
import { ColumnRaw } from "./ColumnRaw";
import { RowRaw } from "./RowRaw";
import { SheetMetaRaw } from "./SheetMetaRaw";
import { SheetConditionalFormatsRaw } from "./SheetRaw/SheetConditionalFormatsRaw";
import { SheetEditProtectionsRaw } from "./SheetRaw/SheetEditProtectionsRaw";
import { SpreadsheetRaw } from "./SpreadsheetRaw";

/**
 * One sheet's grid state by index: rows, columns, table geometry, pruning,
 * queued sheet-level requests, and integrating fetched sheet data into rows,
 * cells and Meta column facts. Conditional format rules and edit protections
 * live in SheetRaw/ and are reached through one-line delegations here.
 * Uniform rows and column facts are SheetMetaRaw; spreadsheet-wide fetch and
 * flush are SpreadsheetRaw. By-name and columnId resolution are Identified/Named.
 */
export class SheetRaw extends SheetCommonRaw {
  get ss(): SpreadsheetRaw {
    return new SpreadsheetRaw(this.spreadsheetRawProps);
  }
  get meta(): SheetMetaRaw {
    return new SheetMetaRaw(this.sheetRawProps);
  }
  private get conditionalFormats(): SheetConditionalFormatsRaw {
    return new SheetConditionalFormatsRaw(this.sheetRawProps);
  }
  private get protections(): SheetEditProtectionsRaw {
    return new SheetEditProtectionsRaw(this.sheetRawProps);
  }
  get rowIndexesAreStale(): boolean {
    return (
      this.sheetState.working.knownTable !== undefined &&
      this.activeTable.rowIndexesAreStale
    );
  }
  get hasFetchedProperties(): boolean {
    return this.sheetState.working.knownTable !== undefined;
  }
  get dataGridRange(): GridRangeProps {
    return {
      sheetId: this.sheetGid,
      startRowIndex: this.schema.topDataRowIdx,
      endRowIndex: this.activeTable.endRowIndex,
      startColumnIndex: this.activeTable.startColumnIndex,
      endColumnIndex: this.activeTable.endColumnIndex,
    };
  }
  get wholeSheetGridRange(): ProtectionGridRange {
    return { sheetId: this.sheetGid };
  }
  rowGridRange(rowIndex: number): GridRangeProps {
    return {
      sheetId: this.sheetGid,
      startRowIndex: rowIndex,
      endRowIndex: rowIndex + 1,
    };
  }
  get title(): string {
    if (this.sheetState.working.title === undefined) {
      throw new Error(
        `Sheet title is null for sheetGid ${this.sheetGid}. Ensure that the sheet properties have been fetched.`,
      );
    }
    return this.sheetState.working.title;
  }
  updateTitle(title: string): this {
    this.updateRequests.updateSheetTitle.push({
      kind: "updateSheetTitle",
      sheetId: this.sheetGid,
      title,
    });
    this.sheetState.working.title = title;
    return this;
  }
  get tables(): TableIdentityRaw[] {
    return this.sheetState.working.tables;
  }
  updateTableName(name: string): this {
    const tableId = this.activeTable.tableId;
    this.updateRequests.updateTableName.push({
      kind: "updateTableName",
      tableId,
      name,
    });
    this._updateWorkingTableName(tableId, name);
    return this;
  }
  get activeRowIndexes(): number[] {
    const indexes = Array.from(this.sheetState.working.rowStates.keys());
    return Arr.sortAscending(indexes);
  }
  get activeRowCount(): number {
    return this.sheetState.working.rowStates.size;
  }
  get lastActiveRowIndex(): number {
    return Math.max(...this.rowStates.keys());
  }
  get rowIndexesActive(): number[] {
    return this.activeRowIndexes.filter((rowIndex) =>
      this.schema.isDataRowIndex(rowIndex),
    );
  }
  get rowIndexesFull(): number[] {
    return Arr.indexesFromUntil(
      this.schema.topDataRowIdx,
      this.activeTable.endRowIndex,
    );
  }
  get rowsFull(): RowRaw[] {
    return this.rowIndexesFull.map((rowIndex) => this.row(rowIndex));
  }
  get rows(): RowRaw[] {
    return this.rowIndexesActive.map((index) => this.row(index));
  }
  get topRow(): RowRaw {
    return this.row(this.schema.topDataRowIdx);
  }
  get rowCount(): number {
    return this.activeRowCount - this.schema.topDataRowIdx;
  }
  // The one place the invariant's threshold is written, so no tier can drift from it.
  get isDownToLastDataRow(): boolean {
    return this.dataRowCountAfterFlush <= 1;
  }
  // Local row state holds only fetched rows, so the table's extent is the source.
  get dataRowCountAfterFlush(): number {
    const { endRowIndex } = this.activeTable;
    return (
      endRowIndex - this.schema.topDataRowIdx - this._queuedRowDeleteCount()
    );
  }
  private _queuedRowDeleteCount(): number {
    let count = 0;
    this.sheetState.writeQueue.rows.forEach((change) => {
      if (change.delete) count++;
    });
    return count;
  }
  get cellStateIsStale(): boolean {
    return this.sheetState.working.cellStateIsStale;
  }
  markRowIndexesStale(): void {
    this.activeTable.markRowIndexesStale();
  }
  invalidateCellState(): void {
    this.sheetState.working.rowStates.clear();
    this.sheetState.working.cellStateIsStale = true;
  }
  findReplace(terms: FindReplaceTerms): this {
    this.ss.findReplace({ ...terms, scope: { sheetId: this.sheetGid } });
    return this;
  }
  clearRowIndexStale(): void {
    this.activeTable.clearRowIndexStale();
  }
  ensureColIndexIsStale(colIndex: number): void {
    this.activeTable.ensureColIndexIsStale(colIndex);
  }
  row(rowIndex: number): RowRaw {
    return new RowRaw({
      rowIndex,
      ...this.sheetRawProps,
    });
  }
  // Every guess this sheet's columns made from a sample had none behind it.
  topDataRowIsBlank(): boolean {
    if (this.topRow.rowIsActive()) {
      return this.fullTableColIndexes.every(
        (colIndex) => this.topRow.valueOrEmpty(colIndex) === "",
      );
    }
    // A queued-delete top row has no cells; column facts still describe the live sheet.
    return this.fullTableColIndexes.every(
      (colIndex) => this.meta.column(colIndex).activeTopValue === "",
    );
  }
  // Either kind of row, for callers that only touch what the two share.
  rowCommon(rowIndex: number): RowCommonRaw {
    if (this.schema.isUniformRowIndex(rowIndex)) {
      return this.meta.uniformRowByIndex(rowIndex);
    } else {
      return this.row(rowIndex);
    }
  }
  column<VN extends CellValueName = CellValueName>(
    colIndex: number,
  ): ColumnRaw<VN> {
    return new ColumnRaw<VN>({
      colIndex,
      ...this.sheetRawProps,
    });
  }
  columnByHeader<VN extends CellValueName = CellValueName>(
    header: string,
  ): ColumnRaw<VN> {
    return this.column<VN>(this.meta.tableHeaderRow.colIndexOfValue(header));
  }
  gatherFetchDataColumnsUsingHeaders<HD extends string>(
    ...headers: HD[]
  ): Record<HD, ColumnRaw> {
    return headers.reduce(
      (acc, header) => {
        acc[header] = this.columnByHeader(header).gatherFetchFull();
        return acc;
      },
      {} as Record<HD, ColumnRaw>,
    );
  }
  gatherFetchProperties(startTableColIndex: number): this {
    // The live start is unknown until this probe comes back, so aim the layout constant.
    this.meta.tableHeaderRow.cell(startTableColIndex).gatherFetchRange();
    return this;
  }
  hasQueuedFullRowFetch(rowIndex: number): boolean {
    return this.sheetState.fetchQueue.toFinalize.rows.has(rowIndex);
  }
  finalizeFetchedCells(): void {
    this.sheetState.fetchQueue.toFinalize.cells.forEach(
      (colIndexes, rowIndex) => {
        const row = this.rowCommon(rowIndex);
        row.ensureStateExists();
        colIndexes.forEach((colIndex) => {
          row.cell(colIndex).ensureActive();
        });
      },
    );
    this.sheetState.fetchQueue.toFinalize.cells.clear();
  }
  // After the backfills above, so a blank fact is sampled rather than built.
  ensureFetchedActiveFacts(): void {
    const { toFinalize } = this.sheetState.fetchQueue;
    if (toFinalize.rows.has(this.schema.topDataRowIdx)) {
      this.meta.ensureTableColumnsActiveFacts();
    }
    toFinalize.columns.forEach((colIndex) => {
      if (!this.isTableColIndex(colIndex)) return;
      this.meta.column(colIndex).ensureActiveFacts();
    });
  }
  integrateSheetState(sheet: SheetSnapshot): void {
    this._initSheetState(sheet);
    this.sheetState.working.cellStateIsStale = false;
    if (sheet.gridBlocks) {
      this._integrateSheetData(sheet.gridBlocks);
    }
  }
  private _integrateSheetData(
    gridBlocks: NonNullable<SheetSnapshot["gridBlocks"]>,
  ): void {
    gridBlocks.forEach((block) => {
      const colIdxBase = block.startColumn;
      block.rows.forEach((rowSnapshot, rowIdxBase) => {
        const rowIndex = rowIdxBase + block.startRow;
        const row = this.rowCommon(rowIndex);
        row.ensureStateExists();
        for (
          let colIdxOffset = 0;
          colIdxOffset < block.columnCount;
          colIdxOffset++
        ) {
          const colIndex = colIdxBase + colIdxOffset;
          const cellData = rowSnapshot.cells[colIdxOffset];
          if (row.rowIsActive()) {
            row.cell(colIndex).integrateSnapshot(cellData);
          }
          if (
            rowIndex === this.schema.topDataRowIdx &&
            this.isTableColIndex(colIndex)
          ) {
            this.meta.column(colIndex).integrateActiveFacts(cellData);
          }
        }
      });
    });
  }
  gatherFetchConditionalFormatRules(): this {
    this.conditionalFormats.gatherFetchConditionalFormatRules();
    return this;
  }
  conditionalFormatRules(): ConditionalFormatRule[] {
    return this.conditionalFormats.conditionalFormatRules();
  }
  addConditionalFormatRule(declaration: ConditionalFormatDeclaration): this {
    this.conditionalFormats.addConditionalFormatRule(declaration);
    return this;
  }
  removeConditionalFormatRules(): this {
    this.conditionalFormats.removeConditionalFormatRules();
    return this;
  }
  addConditionalFormatRuleAt(
    range: GridRangeProps,
    declaration: ConditionalFormatDeclaration,
  ): this {
    this.conditionalFormats.addConditionalFormatRuleAt(range, declaration);
    return this;
  }
  removeConditionalFormatRulesAt(range: GridRangeProps): this {
    this.conditionalFormats.removeConditionalFormatRulesAt(range);
    return this;
  }
  removeConditionalFormatRule(rule: ConditionalFormatRule): this {
    this.conditionalFormats.removeConditionalFormatRule(rule);
    return this;
  }
  markConditionalFormatIndexesStale(): void {
    this.conditionalFormats.markConditionalFormatIndexesStale();
  }
  assertConditionalFormatIndexesNotStale(): void {
    this.conditionalFormats.assertConditionalFormatIndexesNotStale();
  }
  integrateConditionalFormatRules(rules: ConditionalFormatRule[]): void {
    this.conditionalFormats.integrateConditionalFormatRules(rules);
  }
  gatherFetchEditProtections(): this {
    this.protections.gatherFetchEditProtections();
    return this;
  }
  editProtections(): EditProtection[] {
    return this.protections.editProtections();
  }
  addEditWarning(declaration: EditWarningDeclaration = {}): this {
    this.protections.addEditWarning(declaration);
    return this;
  }
  addEditLock(declaration: EditLockDeclaration = {}): this {
    this.protections.addEditLock(declaration);
    return this;
  }
  addEditWarningWholeSheet(
    declaration: WholeSheetEditWarningDeclaration = {},
  ): this {
    this.protections.addEditWarningWholeSheet(declaration);
    return this;
  }
  addEditLockWholeSheet(declaration: WholeSheetEditLockDeclaration = {}): this {
    this.protections.addEditLockWholeSheet(declaration);
    return this;
  }
  addEditWarningAt(
    range: ProtectionGridRange,
    declaration: EditWarningDeclaration = {},
  ): this {
    this.protections.addEditWarningAt(range, declaration);
    return this;
  }
  addEditLockAt(
    range: ProtectionGridRange,
    declaration: EditLockDeclaration = {},
  ): this {
    this.protections.addEditLockAt(range, declaration);
    return this;
  }
  removeEditProtections(): this {
    this.protections.removeEditProtections();
    return this;
  }
  removeEditProtectionsAt(range: ProtectionGridRange): this {
    this.protections.removeEditProtectionsAt(range);
    return this;
  }
  removeEditProtection(protection: EditProtection): this {
    this.protections.removeEditProtection(protection);
    return this;
  }
  removeEditProtectionByDescription(description: string): this {
    this.protections.removeEditProtectionByDescription(description);
    return this;
  }
  removeEditProtectionById(protectionId: number): this {
    this.protections.removeEditProtectionById(protectionId);
    return this;
  }
  markEditProtectionsStale(): void {
    this.protections.markEditProtectionsStale();
  }
  assertEditProtectionsNotStale(): void {
    this.protections.assertEditProtectionsNotStale();
  }
  integrateEditProtections(protections: EditProtection[]): void {
    this.protections.integrateEditProtections(protections);
  }
  // The uniform rows survive, or every later column-index resolution breaks.
  removeRowsExcept(...rowIdxesToKeep: number[]): void {
    const allRowIdxs = Array.from(this.rowStates.keys());
    allRowIdxs.forEach((rowIndex) => {
      if (this.schema.isUniformRowIndex(rowIndex)) return;
      if (!rowIdxesToKeep.includes(rowIndex)) {
        this.rowCommon(rowIndex).remove();
      }
    });
    this.sheetState.working.isPrunedToSelection = true;
  }
  // A whole-column fill ignores active rows, so it would rewrite what a prune excluded.
  validateNotPrunedToSelection(): void {
    if (this.sheetState.working.isPrunedToSelection) {
      throw new Error(
        `Sheet ${this.sheetGid} has been pruned to a selection. A whole-column write would reach the rows the prune excluded.`,
      );
    }
  }
  requestSortGSheet({ colIdxToSortBy, sortOrder }: SortParameters): void {
    this.addSheetChangeToSave({
      action: "sort",
      colIdxToSortBy,
      sortOrder,
    });
  }
  // Value/colour fills stay one repeatCell; a formula fill is pasteData so Sheets parses it.
  gatherFillRequest({
    colIndex,
    startRowIndex,
    endRowIndex,
    formula,
    ...change
  }: ColumnFill): void {
    assertValueAndFormulaExclusive(change.value, formula);
    this.updateRequests.fill.push({
      kind: "fill",
      sheetId: this.sheetGid,
      colIndex,
      startRowIndex,
      endRowIndex,
      ...change,
      ...(formula !== undefined ? { formula } : {}),
    });
  }
  addCheckboxValidationAt(range: BoundedGridRange): this {
    this.activeTable.assertRowIndexesNotStale();
    this.updateRequests.addCheckboxValidation.push({
      kind: "addCheckboxValidation",
      range,
    });
    return this;
  }
  gatherInsertColumnRequest(startColumnIndex: number): void {
    this.updateRequests.insertColumn.push({
      kind: "insertColumn",
      sheetId: this.sheetGid,
      startColumnIndex,
    });
    if (startColumnIndex === this.activeTable.endColumnIndex) {
      this.activeTable.growEndColumnIndex();
    } else {
      this.ensureColIndexIsStale(startColumnIndex);
    }
  }
  gatherColumnTypesRequest(ops: UpdateTableColumnTypeOperation[]): void {
    this._assertColumnTypesUpdateAllowed(ops);
    this.updateRequests.updateTableColumnProperties.push({
      kind: "updateTableColumnProperties",
      tableId: this.activeTable.tableId,
      columnProperties: this._columnTypesColumnProperties(ops),
    });
  }
  markColumnPropertiesStale(): void {
    this.activeTable.markColumnPropertiesStale();
  }
  private _assertColumnTypesUpdateAllowed(
    ops: UpdateTableColumnTypeOperation[],
  ): void {
    const tableId = Val.assert(ops[0], "queued column type").tableId;
    const tableLabel = this._tableLabel(tableId);
    if (
      this.sheetState.working.knownTable === undefined ||
      this.activeTable.tableId !== tableId ||
      ops.some((operation) => operation.tableId !== tableId)
    ) {
      throw new Error(`${tableLabel} is not the fetched Table on that sheet.`);
    }
    const fetched = this.activeTable.columnProperties;
    if (fetched.length === 0) {
      throw new Error(
        `${tableLabel} has no fetched column properties; refetch it before setting a column type.`,
      );
    }
    if (
      this.updateRequests.insertColumn.some(
        ({ sheetId }) => sheetId === this.sheetGid,
      )
    ) {
      throw new Error(
        `Refusing to set column types on ${tableLabel}: the same flush inserts a column on that sheet.`,
      );
    }
    const validated = fetched.filter(
      (column) =>
        column.dataValidationConditionType !== undefined ||
        column.dataValidationValues.length > 0,
    );
    if (validated.length > 0) {
      throw new Error(
        `Refusing to set column types on ${tableLabel}: it would reset the dropdown style and colours on its validated columns ${validated.map(columnLabel).join(", ")}.`,
      );
    }
    ops.forEach(({ columnIndex }) => {
      if (!fetched.some((column) => column.columnIndex === columnIndex)) {
        throw new Error(
          `${tableLabel} has no fetched ${tableColumnLabel(columnIndex)}.`,
        );
      }
    });
  }
  // Full list, since a partial columnProperties replaces the rest.
  private _columnTypesColumnProperties(
    ops: UpdateTableColumnTypeOperation[],
  ): TableColumnPropertiesUpdate[] {
    const typeByIndex = new Map<number, string>(
      ops.map((operation) => [operation.columnIndex, operation.columnType]),
    );
    return this.activeTable.columnProperties.map((column) => {
      const { columnIndex } = column;
      if (column.columnName === undefined) {
        throw new Error(
          `${this._tableLabel(this.activeTable.tableId)} ${tableColumnLabel(columnIndex)} has no columnName; refusing to replace column properties.`,
        );
      }
      const columnType = typeByIndex.get(columnIndex) ?? column.columnType;
      return {
        columnIndex,
        columnName: column.columnName,
        ...(columnType !== undefined ? { columnType } : {}),
      };
    });
  }
  private _tableLabel(tableId: string): string {
    return `Table ${tableId} on "${this.title}"`;
  }
  gatherSortRequest({ colIdxToSortBy, sortOrder }: SortParameters): void {
    this.updateRequests.sort.push({
      kind: "sort",
      sheetId: this.sheetGid,
      startRowIndex: this.schema.topDataRowIdx,
      startColumnIndex: 0,
      colIdxToSortBy,
      sortOrder,
    });
  }
  appendDataRow(): RowRaw {
    const idx = this.activeTable.endRowIndex;
    return this.row(idx).append();
  }
  appendDataRowValues(colValues: Map<number, Value>): RowRaw {
    const row = this.appendDataRow();
    for (const [colIndex, value] of colValues.entries()) {
      row.updateValue(colIndex, value);
    }
    return row;
  }
}

function columnLabel(column: TableColumnSnapshot): string {
  return column.columnName ?? tableColumnLabel(column.columnIndex);
}

function tableColumnLabel(columnIndex: number): string {
  return `table column ${columnIndex}`;
}
