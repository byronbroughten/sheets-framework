import { Arr } from "@byronbroughten/utils/arr";
import { Val } from "@byronbroughten/utils/val";

import type {
  CellValue,
  CellValueName,
} from "../00_Source/CellValues/cellValues";
import type { ConditionalFormatDeclaration } from "../00_Source/RawSource/ConditionalFormat";
import type {
  EditLockDeclaration,
  EditWarningDeclaration,
} from "../00_Source/RawSource/EditProtection";
import type {
  BoundedGridRange,
  FindReplaceTerms,
  GridBlockSnapshot,
  GridCellSnapshot,
  InsertRangeOperation,
} from "../00_Source/RawSource/RawSource";
import {
  type SheetColIndex,
  SheetIndex,
  type SheetRowIndex,
} from "../00_Source/RawSource/SheetIndex";
import { type HeadRole, headRows } from "../01_SpreadsheetSchema/headRows";
import { assertValueAndFormulaExclusive } from "./CellRaw";
import type { RowCommonRaw } from "./ClassBases/RowCommonRaw";
import { originOf } from "./ClassBases/TableBaseRaw";
import { TableCommonRaw } from "./ClassBases/TableCommonRaw";
import type {
  ChangeTableColumns,
  ColumnFill,
  ColumnMoveSpan,
  NewColumnHeadCells,
  SortParameters,
  TableCell,
  TableFindReplace,
  TableGridRange,
} from "./ClassTypes/StateRaw";
import { ColumnRaw } from "./ColumnRaw";
import { HeadRowRaw } from "./HeadRowRaw";
import { RowRaw } from "./RowRaw";
import { SheetRaw } from "./SheetRaw";
import { SpreadsheetRaw } from "./SpreadsheetRaw";
import { TableProfileRaw } from "./TableProfileRaw";
import { TableColumnQueuerRaw } from "./TableRaw/TableColumnQueuerRaw";
import { TableColumnResolverRaw } from "./TableRaw/TableColumnResolverRaw";
import { TableColumnTypeGathererRaw } from "./TableRaw/TableColumnTypeGathererRaw";
import { TableConditionalFormatAndProtectionQueuerRaw } from "./TableRaw/TableConditionalFormatAndProtectionQueuerRaw";
import { columnRun, TableRowAppenderRaw } from "./TableRaw/TableRowAppenderRaw";

/**
 * One Table's state by Table-relative index: rows, columns, pruning, queued
 * Table-level requests, and integrating fetched cells into its rows, cells and
 * sampled column facts. Its sheet's title, grid size, conditional format rules
 * and edit protections are SheetRaw, reached as `table.sheet`.
 * Descriptive facts are TableProfileRaw and ColumnProfileRaw; column ID lookups, column
 * changes, column types, and conditional format rules and protections are TableRaw/;
 * spreadsheet-wide fetch and flush are SpreadsheetRaw. Config-aware lookups and names are Identified/Named.
 */
export class TableRaw extends TableCommonRaw {
  get ss(): SpreadsheetRaw {
    return new SpreadsheetRaw(this.spreadsheetRawProps);
  }
  get sheet(): SheetRaw {
    return new SheetRaw({
      ...this.spreadsheetRawProps,
      sheetGid: this.sheetGid,
    });
  }
  get profile(): TableProfileRaw {
    return new TableProfileRaw(this.tableRawProps);
  }
  get columnResolver(): TableColumnResolverRaw {
    return new TableColumnResolverRaw(this.tableRawProps);
  }
  private get columnQueuer(): TableColumnQueuerRaw {
    return new TableColumnQueuerRaw(this.tableRawProps);
  }
  private get columnTypeGatherer(): TableColumnTypeGathererRaw {
    return new TableColumnTypeGathererRaw(this.tableRawProps);
  }
  private get rowAppender(): TableRowAppenderRaw {
    return new TableRowAppenderRaw(this.tableRawProps);
  }
  get conditionalFormatAndProtectionQueuer(): TableConditionalFormatAndProtectionQueuerRaw {
    return new TableConditionalFormatAndProtectionQueuerRaw(this.tableRawProps);
  }
  get hasFetchedProperties(): boolean {
    return this.tableProperties !== undefined;
  }
  get dataTableGridRange(): TableGridRange {
    return {
      startRowIndex: 0,
      endRowIndex: this.dataRowCount,
      startColIndex: 0,
      endColIndex: this.columnCount,
    };
  }
  // The live Table's columns only, so a range built on it never reaches a neighbour.
  dataRowGridRange(rowIndex: number): BoundedGridRange {
    return this._boundedSheetGridRange(
      {
        startRowIndex: rowIndex,
        endRowIndex: rowIndex + 1,
        startColIndex: 0,
        endColIndex: this.columnCount,
      },
      this.originAtGathering(),
    );
  }
  updateTableName(name: string): this {
    const tableId = this.tableId;
    this.writeOperations.renameTable.push({
      kind: "renameTable",
      tableId,
      name,
    });
    this._updateWorkingTableName(tableId, name);
    return this;
  }
  // The head rows' range is gathered, so it takes the stale guard of any write placed by row.
  queueDelete(): this {
    this.assertRowIndexesNotStale();
    return this.queueTableWrite({ action: "deleteTable" });
  }
  updateHeader(columnId: string, header: string): this {
    this.headRow("header").updateValue(
      this.columnResolver.colIndexOf(columnId),
      header,
    );
    return this;
  }
  get workingRowIndexesWithHead(): number[] {
    const indexes = Array.from(this.rowStates.keys());
    return Arr.sortAscending(indexes);
  }
  get workingRowCount(): number {
    return this.rowStates.size;
  }
  get lastWorkingRowIndex(): number {
    return Math.max(...this.rowStates.keys());
  }
  get workingRowIndexes(): number[] {
    return this.workingRowIndexesWithHead.filter((rowIndex) => rowIndex >= 0);
  }
  get rowIndexesFull(): number[] {
    return Arr.indexesFromUntil(0, this.dataRowCount);
  }
  get rowsFull(): RowRaw[] {
    return this.rowIndexesFull.map((rowIndex) => this.row(rowIndex));
  }
  get rows(): RowRaw[] {
    return this.workingRowIndexes.map((index) => this.row(index));
  }
  get topRow(): RowRaw {
    return this.row(0);
  }
  get rowCount(): number {
    return this.workingRowIndexes.length;
  }
  // The one place the invariant's threshold is written, so no tier can drift from it.
  get isDownToLastDataRow(): boolean {
    return this.dataRowCountAfterFlush <= 1;
  }
  // Local row state holds only fetched rows, so the table's extent is the source.
  get dataRowCountAfterFlush(): number {
    return this.dataRowCount - this._queuedRowDeleteCount();
  }
  private _queuedRowDeleteCount(): number {
    let count = 0;
    this.tableState.writeQueue.rows.forEach((writes) => {
      if (writes.deleteRow) count++;
    });
    return count;
  }
  get cellStateIsStale(): boolean {
    return this.tableState.working.cellStateIsStale;
  }
  invalidateCellState(): void {
    this.rowStates.clear();
    this.tableState.working.cellStateIsStale = true;
  }
  // Reaches every body row like a whole-column fill, so it takes the same guards.
  findReplace(terms: FindReplaceTerms): this {
    this.assertRowIndexesNotStale();
    this.validateNotPrunedToSelection();
    return this.queueTableWrite({
      action: "findReplace",
      terms,
      startColIndex: 0,
      endColIndex: this.columnCount,
    });
  }
  appendColumn(headCells: NewColumnHeadCells): number {
    return this.columnQueuer.appendColumn(headCells);
  }
  insertColumn(colIndex: number, headCells: NewColumnHeadCells): this {
    this.columnQueuer.insertColumn(colIndex, headCells);
    return this;
  }
  deleteColumn(colIndex: number): this {
    this.columnQueuer.deleteColumn(colIndex);
    return this;
  }
  moveColumns(span: ColumnMoveSpan): this {
    this.columnQueuer.moveColumns(span);
    return this;
  }
  row(rowIndex: number): RowRaw {
    return new RowRaw({
      rowIndex,
      ...this.tableRawProps,
    });
  }
  // Every guess this sheet's columns made from a sample had none behind it.
  topDataRowIsBlank(): boolean {
    if (this.topRow.rowInWorking()) {
      return this.fullTableColIndexes.every(
        (colIndex) => this.topRow.valueOrEmpty(colIndex) === "",
      );
    }
    // A queued-delete top row has no cells; column facts still describe the live sheet.
    return this.fullTableColIndexes.every(
      (colIndex) => this.column(colIndex).profile.topValue === "",
    );
  }
  headRow<HR extends HeadRole>(headRole: HR): HeadRowRaw<HR> {
    return new HeadRowRaw({ ...this.tableRawProps, headRole });
  }
  headRowByIndex(rowIndex: number): HeadRowRaw {
    return this.headRow(headRows.rolesAt(rowIndex)[0]);
  }
  // Either kind of row, for callers that only touch what the two share.
  rowCommon(rowIndex: number): RowCommonRaw {
    if (headRows.isIndex(rowIndex)) {
      return this.headRowByIndex(rowIndex);
    } else {
      return this.row(rowIndex);
    }
  }
  column<VN extends CellValueName = CellValueName>(
    colIndex: number,
  ): ColumnRaw<VN> {
    return new ColumnRaw<VN>({
      colIndex,
      ...this.tableRawProps,
    });
  }
  columnByHeader<VN extends CellValueName = CellValueName>(
    header: string,
  ): ColumnRaw<VN> {
    return this.column<VN>(this.columnResolver.colIndexOfHeader(header));
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
  gatherFetchProperties(): this {
    this.sheet.gatherFetchHeaderZone();
    this._prepFetchHeadRowBackfills();
    return this;
  }
  // The zone holds every head row, so a Table it brought finalizes them with no fetch of its own.
  integrateHeaderZone(): void {
    this._prepFetchHeadRowBackfills();
    this.tableState.fetchQueue.toFinalize.rows.add(this.schema.colIdRowIndex);
    this.finalizeFetches();
  }
  private _prepFetchHeadRowBackfills(): void {
    headRows.indexes().forEach((rowIndex) => {
      this.headRowByIndex(rowIndex).cell(0).prepFetchBackfill();
    });
  }
  holdsActionCellAt(
    sheetRowIndex: SheetRowIndex,
    sheetColIndex: SheetColIndex,
  ): boolean {
    if (!this.hasFetchedProperties) return false;
    const { origin } = this;
    return (
      sheetRowIndex === origin.headSheetRowIndex("action") &&
      this.isTableColIndex(origin.colIndex(sheetColIndex))
    );
  }
  hasQueuedFullRowFetch(rowIndex: number): boolean {
    return this.tableState.fetchQueue.toFinalize.rows.has(rowIndex);
  }
  // Backfills every range fetched this cycle, since Sheets omits empty cells and whole blank rows.
  finalizeFetches(): void {
    const { toFinalize } = this.tableState.fetchQueue;
    this._finalizeFetchedCells();
    if (toFinalize.rows.size === 0 && toFinalize.columns.size === 0) return;
    if (toFinalize.rows.has(this.schema.colIdRowIndex)) {
      this.tableState.working.hasFetchedColumnIds = true;
    }
    toFinalize.rows.forEach((rowIndex) => {
      this.rowCommon(rowIndex).ensureFullWorkingDataCells();
    });
    toFinalize.columns.forEach((colIndex) => {
      this.column(colIndex).ensureFullWorkingDataCells();
    });
    this._ensureFetchedSampledFacts();
    toFinalize.rows.clear();
    toFinalize.columns.clear();
  }
  private _finalizeFetchedCells(): void {
    this.tableState.fetchQueue.toFinalize.cells.forEach(
      (colIndexes, rowIndex) => {
        const row = this.rowCommon(rowIndex);
        row.ensureStateExists();
        colIndexes.forEach((colIndex) => {
          row.cell(colIndex).ensureInWorking();
        });
      },
    );
    this.tableState.fetchQueue.toFinalize.cells.clear();
  }
  // After the backfills above, so a blank fact is sampled rather than built.
  private _ensureFetchedSampledFacts(): void {
    const { toFinalize } = this.tableState.fetchQueue;
    // Only table columns: a fact is always reached through a column ID.
    if (toFinalize.rows.has(0)) {
      this.fullTableColIndexes.forEach((colIndex) => {
        this._ensureSampledFacts(colIndex);
      });
    }
    toFinalize.columns.forEach((colIndex) => {
      if (!this.isTableColIndex(colIndex)) return;
      this._ensureSampledFacts(colIndex);
    });
  }
  // Gap-filling only, so a fact the payload described always wins.
  private _ensureSampledFacts(colIndex: number): void {
    if (this.columnStates.get(colIndex)?.sampledFacts !== undefined) return;
    if (!this.column(colIndex).topCell.inWorking) return; // no top data row to sample
    this._integrateSampledFacts(colIndex, undefined);
  }
  private _integrateSampledFacts(
    colIndex: number,
    cell: GridCellSnapshot | undefined,
  ): void {
    this._ensureColumnState(colIndex).sampledFacts = {
      isFormula: cell?.isFormula ?? false,
      numberFormatType: cell?.numberFormatType,
      dataValidationConditionType: cell?.dataValidationConditionType,
      topValue: cell?.value ?? "", // from the payload, so a deleted top data row still describes the column
      topFormula: cell?.formula,
    };
  }
  // Each Table takes only the cells inside its own area, so a loose cell reaches no state.
  integrateGridBlocks(gridBlocks: GridBlockSnapshot[]): void {
    this.tableState.working.cellStateIsStale = false;
    const properties = Val.assert(
      this.tableProperties,
      `Table properties for ${this.sheetLabel}`,
    );
    const origin = originOf(properties);
    const bodyRowCount = properties.endRowIndex - properties.startRowIndex - 1;
    const columnCount = properties.endColumnIndex - properties.startColumnIndex;
    gridBlocks.forEach((block) => {
      const firstColIndex = origin.colIndex(block.startColumn);
      const firstRowIndex = origin.rowIndex(block.startRow);
      block.rows.forEach((rowSnapshot, rowOffset) => {
        const rowIndex = firstRowIndex + rowOffset;
        // The column ID row is the topmost head row.
        if (rowIndex < this.schema.colIdRowIndex) return;
        if (rowIndex >= bodyRowCount) return;
        const row = this.rowCommon(rowIndex);
        row.ensureStateExists();
        for (
          let colIdxOffset = 0;
          colIdxOffset < block.columnCount;
          colIdxOffset++
        ) {
          const colIndex = firstColIndex + colIdxOffset;
          if (colIndex < 0 || colIndex >= columnCount) continue;
          const cellData = rowSnapshot.cells[colIdxOffset];
          if (row.rowInWorking()) {
            row.cell(colIndex).integrateSnapshot(cellData);
          }
          if (rowIndex === 0) {
            this._integrateSampledFacts(colIndex, cellData);
          }
        }
      });
    });
  }
  addConditionalFormatRule(declaration: ConditionalFormatDeclaration): this {
    return this.addConditionalFormatRuleAt(
      this.dataTableGridRange,
      declaration,
    );
  }
  removeConditionalFormatRules(): this {
    return this.removeConditionalFormatRulesAt(this.dataTableGridRange);
  }
  addEditWarning(declaration: EditWarningDeclaration = {}): this {
    return this.addEditWarningAt(this.dataTableGridRange, declaration);
  }
  addEditLock(declaration: EditLockDeclaration = {}): this {
    return this.addEditLockAt(this.dataTableGridRange, declaration);
  }
  removeEditProtections(): this {
    return this.removeEditProtectionsAt(this.dataTableGridRange);
  }
  addConditionalFormatRuleAt(
    range: TableGridRange,
    declaration: ConditionalFormatDeclaration,
  ): this {
    this.conditionalFormatAndProtectionQueuer.addConditionalFormatRuleAt(
      range,
      declaration,
    );
    return this;
  }
  removeConditionalFormatRulesAt(range: TableGridRange): this {
    this.conditionalFormatAndProtectionQueuer.removeConditionalFormatRulesAt(
      range,
    );
    return this;
  }
  addEditWarningAt(
    range: TableGridRange,
    declaration: EditWarningDeclaration = {},
  ): this {
    this.conditionalFormatAndProtectionQueuer.addEditWarningAt(
      range,
      declaration,
    );
    return this;
  }
  addEditLockAt(
    range: TableGridRange,
    declaration: EditLockDeclaration = {},
  ): this {
    this.conditionalFormatAndProtectionQueuer.addEditLockAt(range, declaration);
    return this;
  }
  removeEditProtectionsAt(range: TableGridRange): this {
    this.conditionalFormatAndProtectionQueuer.removeEditProtectionsAt(range);
    return this;
  }
  // The head rows survive, or every later column-index resolution breaks.
  removeRowsExcept(...rowIdxesToKeep: number[]): void {
    const allRowIdxs = Array.from(this.rowStates.keys());
    allRowIdxs.forEach((rowIndex) => {
      if (headRows.isIndex(rowIndex)) return;
      if (!rowIdxesToKeep.includes(rowIndex)) {
        this.rowCommon(rowIndex).remove();
      }
    });
    this.tableState.working.isPrunedToSelection = true;
  }
  // A whole-column fill ignores working rows, so it would rewrite what a prune excluded.
  validateNotPrunedToSelection(): void {
    if (this.tableState.working.isPrunedToSelection) {
      throw new Error(
        `Sheet ${this.sheetGid} has been pruned to a selection. A whole-column write would reach the rows the prune excluded.`,
      );
    }
  }
  queueSort({ colIdxToSortBy, sortOrder }: SortParameters): void {
    this.queueTableWrite({
      action: "sort",
      colIdxToSortBy,
      sortOrder,
    });
  }
  // Value/colour fills stay one repeatCell; a formula fill is pasteData so Sheets parses it.
  gatherFillColumnOperation({
    colIndex,
    startRowIndex,
    endRowIndex,
    formula,
    ...change
  }: ColumnFill): void {
    assertValueAndFormulaExclusive(change.value, formula);
    const origin = this.originAtGathering();
    this.writeOperations.fillColumn.push({
      kind: "fillColumn",
      sheetId: this.sheetGid,
      colIndex: origin.sheetColIndex(colIndex),
      startRowIndex: origin.sheetRowIndex(startRowIndex),
      endRowIndex: origin.sheetRowIndex(endRowIndex),
      ...change,
      ...(formula !== undefined ? { formula } : {}),
    });
  }
  addCheckboxValidationAt({ rowIndex, colIndex }: TableCell): this {
    this.assertRowIndexesNotStale();
    return this.queueTableWrite({
      action: "addCheckboxValidation",
      rowIndex,
      colIndex,
    });
  }
  gatherQueuedTableWrites(): void {
    const { writes } = this;
    writes.fillColumns.forEach((fill) => {
      this.gatherFillColumnOperation(fill);
    });
    writes.findReplaces.forEach((findReplace) => {
      this.gatherFindReplaceOperation(findReplace);
    });
    writes.checkboxCells.forEach((cell) => {
      this.gatherCheckboxValidationOperation(cell);
    });
    const { conditionalFormatAndProtectionQueuer } = this;
    writes.conditionalFormatRules.forEach((rule) => {
      conditionalFormatAndProtectionQueuer.gatherConditionalFormatRuleOperation(
        rule,
      );
    });
    writes.editProtections.forEach((protection) => {
      conditionalFormatAndProtectionQueuer.gatherEditProtectionOperation(
        protection,
      );
    });
    if (writes.deletesTable) this._gatherDeleteTableOperation();
  }
  // Before the row deletes, which are sent after it, so the head rows are where it measures them.
  private _gatherDeleteTableOperation(): void {
    this.writeOperations.deleteTable.push({
      kind: "deleteTable",
      tableId: this.tableId,
      headRows: this._boundedSheetGridRange(
        {
          startRowIndex: headRows.topIndex,
          endRowIndex: headRows.index("header"),
          startColIndex: 0,
          endColIndex: this.columnCount,
        },
        this.originAtGathering(),
      ),
    });
  }
  gatherCheckboxValidationOperation({ rowIndex, colIndex }: TableCell): void {
    const range = this._boundedSheetGridRange(
      {
        startRowIndex: rowIndex,
        endRowIndex: rowIndex + 1,
        startColIndex: colIndex,
        endColIndex: colIndex + 1,
      },
      this.originAtGathering(),
    );
    this.writeOperations.addCheckboxValidation.push({
      kind: "addCheckboxValidation",
      range,
    });
  }
  // Sent before the row deletes, so it spans the body as it stands before them.
  gatherFindReplaceOperation({
    terms,
    startColIndex,
    endColIndex,
  }: TableFindReplace): void {
    const body = this._bodyGridRangeAtGathering(this.dataRowCount);
    this.writeOperations.findReplace.push({
      kind: "findReplace",
      terms,
      scope: { range: columnRun(body, startColIndex, endColIndex) },
    });
  }
  columnInsertSplitting(neighbourName: string): string {
    return this.columnQueuer.columnInsertSplitting(neighbourName);
  }
  columnChangeSplitting(
    change: ChangeTableColumns,
    insert: InsertRangeOperation,
    neighbourName: string,
  ): string {
    return this.columnQueuer.columnChangeSplitting(
      change,
      insert,
      neighbourName,
    );
  }
  get columnInsertLabel(): string {
    return this.columnQueuer.columnInsertLabel;
  }
  get hasQueuedColumnInsert(): boolean {
    return this.writes.insertTableEndColumnCount > 0;
  }
  get hasQueuedColumnWrites(): boolean {
    return this.hasQueuedColumnInsert || this.writes.columnChanges.length > 0;
  }
  gatherInsertTableEndColumnsOperation(): void {
    this.columnQueuer.gatherInsertTableEndColumnsOperation();
  }
  gatherColumnChangesOperation(): void {
    this.columnQueuer.gatherColumnChangesOperation();
  }
  gatherAppendTableRowsOperation(): void {
    this.rowAppender.gatherAppendTableRowsOperation();
  }
  growWithGridEdge(addedRowCount: number): void {
    this.rowAppender.growWithGridEdge(addedRowCount);
  }
  assertDropdownStyleResetAllowed(): void {
    this.columnTypeGatherer.assertDropdownStyleResetAllowed();
  }
  gatherSetTableColumnPropertiesOperation(): void {
    this.columnTypeGatherer.gatherSetTableColumnPropertiesOperation();
  }
  // Sent after the row deletes, so it spans only the body rows they leave.
  gatherSortTableOperation(): void {
    const { sort } = this.writes;
    if (sort === undefined) return;
    const { colIdxToSortBy, sortOrder } = sort;
    const body = this._bodyGridRangeAtGathering(this.dataRowCountAfterFlush);
    this.writeOperations.sortTable.push({
      kind: "sortTable",
      range: body,
      colIdxToSortBy: SheetIndex.col(body.startColumnIndex + colIdxToSortBy),
      sortOrder,
    });
  }
  private _bodyGridRangeAtGathering(dataRowCount: number): BoundedGridRange {
    return this._boundedSheetGridRange(
      {
        startRowIndex: 0,
        endRowIndex: dataRowCount,
        startColIndex: 0,
        endColIndex: this.columnCount,
      },
      this.originAtGathering(),
    );
  }
  // A head row's range spans every Table on the sheet.
  get tablesOnSheet(): TableRaw[] {
    return this.tableIds().map((tableId) => this.ss.table(tableId));
  }
  get neighbourTables(): TableRaw[] {
    return this.tableIds()
      .filter((tableId) => tableId !== this.tableId)
      .map((tableId) => this.ss.table(tableId));
  }
  // By geometry alone: the column ID row may give way to metadata on the header row.
  get neighboursUnderHeadRows(): TableRaw[] {
    return this.neighbourTables.filter((other) => this.headRowsSitOn(other));
  }
  // Growth copies the named formula columns from the model row; Raw can't tell which they are.
  appendDataRow(formulaColIndexes: readonly number[] = []): RowRaw {
    const row = this.row(this.dataRowCount).append();
    this.queueTableWrite({
      action: "copyAppendedFormulas",
      colIndexes: formulaColIndexes,
    });
    return row;
  }
  appendDataRowValues(colValues: Map<number, CellValue>): RowRaw {
    const row = this.appendDataRow();
    for (const [colIndex, value] of colValues.entries()) {
      row.updateValue(colIndex, value);
    }
    return row;
  }
}
