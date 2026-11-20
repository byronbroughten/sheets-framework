import { Arr } from "@byronbroughten/utils/arr";
import { Val } from "@byronbroughten/utils/val";

import type {
  CellValue,
  CellValueName,
} from "../00_Source/CellValues/cellValues";
import type {
  ConditionalFormatDeclaration,
  ModelableConditionalFormatRule,
} from "../00_Source/RawSource/ConditionalFormat";
import type {
  EditLockDeclaration,
  EditProtectionContent,
  EditWarningDeclaration,
} from "../00_Source/RawSource/EditProtection";
import type {
  BoundedGridRange,
  CopyPasteOperation,
  CopyPasteType,
  FindReplaceTerms,
  GridBlockSnapshot,
  GridCellSnapshot,
  GridRangeProps,
  InsertRangeOperation,
  TableColumnPropertiesUpdate,
  TableColumnSnapshot,
  TableColumnValidationRule,
} from "../00_Source/RawSource/RawSource";
import {
  type SheetColIndex,
  SheetIndex,
  type SheetRowIndex,
} from "../00_Source/RawSource/SheetIndex";
import { type HeadRole, headRows } from "../01_SpreadsheetSchema/headRows";
import type { TableOrigin } from "../01_SpreadsheetSchema/TableOrigin";
import { assertValueAndFormulaExclusive } from "./CellRaw";
import type { RowCommonRaw } from "./ClassBases/RowCommonRaw";
import { hasQueuedWrites, originOf } from "./ClassBases/TableBaseRaw";
import { columnCountChange, TableCommonRaw } from "./ClassBases/TableCommonRaw";
import { queuedColumnType } from "./ClassTypes/queuedColumnType";
import {
  type ChangeTableColumns,
  type ColumnFill,
  type ColumnMoveSpan,
  type NewColumnHeadCells,
  type SortParameters,
  type TableCell,
  type TableColumnChange,
  type TableColumnChangeOperation,
  type TableConditionalFormatRule,
  type TableEditProtection,
  type TableFindReplace,
  type TableGridRange,
  type TableWrites,
} from "./ClassTypes/StateRaw";
import { ColumnRaw } from "./ColumnRaw";
import { HeadRowRaw } from "./HeadRowRaw";
import { RowRaw } from "./RowRaw";
import { SheetRaw } from "./SheetRaw";
import { SpreadsheetRaw } from "./SpreadsheetRaw";
import { TableProfileRaw } from "./TableProfileRaw";
import { TableColumnResolverRaw } from "./TableRaw/TableColumnResolverRaw";

const columnChangeRequests = {
  insertColumn: { kind: "insertRange", action: "Inserting a column inside" },
  deleteColumn: { kind: "deleteRange", action: "Deleting a column of" },
  moveColumns: { kind: "cutPaste", action: "Moving a column of" },
} as const satisfies Record<
  TableColumnChange["kind"],
  { kind: TableColumnChangeOperation["kind"]; action: string }
>;

type ColumnCountChange = Exclude<TableColumnChange, { kind: "moveColumns" }>;
type ColumnMove = Extract<TableColumnChange, { kind: "moveColumns" }>;

/**
 * One Table's state by Table-relative index: rows, columns, pruning, queued
 * Table-level requests, and integrating fetched cells into its rows, cells and
 * sampled column facts. Its sheet's title, grid size, conditional format rules
 * and edit protections are SheetRaw, reached as `table.sheet`.
 * Descriptive facts are TableProfileRaw and ColumnProfileRaw; column ID lookups are TableRaw/;
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
  // Past the inserts already queued, since each lands at the Table end as it stands then.
  appendColumn(headCells: NewColumnHeadCells): number {
    const colIndex = this.columnCount + this.writes.insertTableEndColumnCount;
    this.queueTableWrite({ action: "insertTableEndColumn" });
    this.column(colIndex)._initHeadCells(headCells);
    return colIndex;
  }
  // Strictly inside, so the band grows the Table with no range update; appendColumn adds at the end.
  insertColumn(colIndex: number, headCells: NewColumnHeadCells): this {
    this.assertRowIndexesNotStale();
    this._assertStrictlyInside(
      colIndex,
      `a column inserted at ${tableColumnLabel(colIndex)}`,
    );
    return this.queueTableWrite({
      action: "insertColumn",
      colIndex,
      headCells,
    });
  }
  deleteColumn(colIndex: number): this {
    this.assertRowIndexesNotStale();
    const columnCount = this._columnCountAfterChanges();
    if (
      !Number.isInteger(colIndex) ||
      colIndex < 0 ||
      colIndex >= columnCount
    ) {
      throw new Error(
        `${this.tableLabel} has no ${tableColumnLabel(colIndex)} to delete: it has ${columnCount} columns.`,
      );
    }
    if (columnCount === 1) {
      throw new Error(
        `Refusing to delete the last column of ${this.tableLabel}: a Table keeps at least one.`,
      );
    }
    return this.queueTableWrite({ action: "deleteColumn", colIndex });
  }
  // Lands strictly inside, like insertColumn: an insert at the Table's first column lands outside it (sheets-behaviour.md).
  moveColumns({ colIndex, count, beforeColIndex }: ColumnMoveSpan): this {
    this.assertRowIndexesNotStale();
    const columnCount = this._columnCountAfterChanges();
    const isBlockInside =
      Number.isInteger(colIndex) &&
      Number.isInteger(count) &&
      count >= 1 &&
      colIndex >= 0 &&
      colIndex + count <= columnCount;
    if (!isBlockInside) {
      throw new Error(
        `${this.tableLabel} has no ${count} column(s) from ${tableColumnLabel(colIndex)} to move: it has ${columnCount} columns.`,
      );
    }
    this._assertStrictlyInside(
      beforeColIndex,
      `columns moved before ${tableColumnLabel(beforeColIndex)}`,
    );
    if (colIndex <= beforeColIndex && beforeColIndex <= colIndex + count) {
      throw new Error(
        `${this.tableLabel} already has its ${count} column(s) from ${tableColumnLabel(colIndex)} before ${tableColumnLabel(beforeColIndex)}.`,
      );
    }
    return this.queueTableWrite({
      action: "moveColumns",
      colIndex,
      count,
      beforeColIndex,
    });
  }
  private _assertStrictlyInside(colIndex: number, taken: string): void {
    const columnCount = this._columnCountAfterChanges();
    if (Number.isInteger(colIndex) && colIndex >= 1 && colIndex < columnCount) {
      return;
    }
    throw new Error(
      `${this.tableLabel} can't take ${taken}: it must sit strictly inside its ${columnCount} columns, so from 1 to ${columnCount - 1}.`,
    );
  }
  private _columnCountAfterChanges(): number {
    return this.writes.columnChanges.reduce(
      (count, { kind }) =>
        count + columnCountChange[columnChangeRequests[kind].kind],
      this.columnCount,
    );
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
  // Checked against what the sheet holds now; converted again at gathering, after this batch's shifts.
  addConditionalFormatRuleAt(
    range: TableGridRange,
    declaration: ConditionalFormatDeclaration,
  ): this {
    this.sheet.assertRowIndexesNotStale();
    this.sheet.assertConditionalFormatIndexesNotStale();
    const rule = this._sheetConditionalFormatRule(
      { range, ...declaration },
      this.tableOrigin(),
    );
    const queued = this.tablesOnSheet.flatMap((table) =>
      table._queuedSheetConditionalFormatRules(),
    );
    if (this.sheet.hasPendingConditionalFormatRule(rule, queued)) return this;
    return this.queueTableWrite({
      action: "addConditionalFormatRule",
      range,
      ...declaration,
    });
  }
  removeConditionalFormatRulesAt(range: TableGridRange): this {
    this.sheet.removeConditionalFormatRulesAt(
      this._sheetGridRange(range, this.tableOrigin()),
    );
    return this;
  }
  addEditWarningAt(
    range: TableGridRange,
    declaration: EditWarningDeclaration = {},
  ): this {
    return this._queueEditProtection({
      kind: "warning",
      range,
      description: declaration.description ?? "",
      users: [],
      groups: [],
      unprotectedRanges: [],
    });
  }
  addEditLockAt(
    range: TableGridRange,
    declaration: EditLockDeclaration = {},
  ): this {
    return this._queueEditProtection({
      kind: "lock",
      range,
      description: declaration.description ?? "",
      users: declaration.users ?? [],
      groups: declaration.groups ?? [],
      unprotectedRanges: [],
    });
  }
  removeEditProtectionsAt(range: TableGridRange): this {
    this.sheet.removeEditProtectionsAt(
      this._sheetGridRange(range, this.tableOrigin()),
    );
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
    writes.conditionalFormatRules.forEach((rule) => {
      this.gatherConditionalFormatRuleOperation(rule);
    });
    writes.editProtections.forEach((protection) => {
      this.gatherEditProtectionOperation(protection);
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
  gatherConditionalFormatRuleOperation(rule: TableConditionalFormatRule): void {
    this.writeOperations.addConditionalFormatRule.push({
      kind: "addConditionalFormatRule",
      index: 0,
      rule: this._sheetConditionalFormatRule(rule, this.originAtGathering()),
    });
  }
  gatherEditProtectionOperation(protection: TableEditProtection): void {
    let origin = this.tableOrigin();
    if (hasTableRows(protection.range)) origin = this.originAtGathering();
    this.writeOperations.addProtectedRange.push({
      kind: "addProtectedRange",
      protection: this._sheetEditProtection(protection, origin),
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
    return splittingPhrase(this.columnInsertLabel, neighbourName);
  }
  // A move's gap is the insert its cutPaste follows, as _columnMoveOperations orders them.
  columnChangeSplitting(
    { operations }: ChangeTableColumns,
    insert: InsertRangeOperation,
    neighbourName: string,
  ): string {
    const next = operations[operations.indexOf(insert) + 1];
    const changeKind =
      next?.kind === "cutPaste" ? "moveColumns" : "insertColumn";
    return splittingPhrase(this._columnChangeLabel(changeKind), neighbourName);
  }
  get columnInsertLabel(): string {
    return `Inserting a column at the end of ${this._namedTableOnSheet}`;
  }
  get hasQueuedColumnInsert(): boolean {
    return this.writes.insertTableEndColumnCount > 0;
  }
  get hasQueuedColumnWrites(): boolean {
    return this.hasQueuedColumnInsert || this.writes.columnChanges.length > 0;
  }
  // From the top head row down, so the head rows move with the Table; the widen is what grows it.
  gatherInsertTableEndColumnsOperation(): void {
    const insertCount = this.writes.insertTableEndColumnCount;
    if (insertCount === 0) return;
    const { columnCount } = this;
    const newColumns = this._headAndBodyBand(
      columnCount,
      columnCount + insertCount,
    );
    this._validateNoNeighbourSplitBy(newColumns, this.columnInsertLabel);
    this.sheet.queueGridColumnsThrough(newColumns.endColumnIndex);
    this.writeOperations.insertTableEndColumns.push({
      tableId: this.tableId,
      newColumns,
      operations: [
        { kind: "insertRange", range: newColumns, shiftDimension: "COLUMNS" },
        {
          kind: "updateTableRange",
          tableId: this.tableId,
          range: {
            ...newColumns,
            startRowIndex: this.startRowIndex,
            startColumnIndex: this.startColumnIndex,
          },
        },
      ],
    });
    this.growColumnCount(insertCount);
  }
  private _validateNoNeighbourSplitBy(
    band: BoundedGridRange,
    actionLabel: string,
  ): void {
    const split = this.neighbourTables.find((neighbour) =>
      neighbour.isSplitBy(band),
    );
    if (split === undefined) return;
    throw new Error(
      `${splittingPhrase(actionLabel, split.name)} with its head rows. Move "${split.name}" so that it and its head rows sit entirely within rows ${band.startRowIndex + 1}–${band.endRowIndex}, or entirely outside them.`,
    );
  }
  // Sent in a flush of its own on this Table, since every other queued write is indexed by the columns before it.
  gatherColumnChangesOperation(): void {
    const { columnChanges } = this.writes;
    if (columnChanges.length === 0) return;
    this._assertColumnChangesAlone();
    const startColumnIndex = this.originAtGathering().sheetColIndex(0);
    const operations = columnChanges.flatMap((change) => {
      if (change.kind === "moveColumns") {
        return this._columnMoveOperations(change);
      }
      return [this._columnChangeOperation(change)];
    });
    this.markColumnPropertiesStale();
    this.writeOperations.changeTableColumns.push({
      tableId: this.tableId,
      sheetId: this.sheetGid,
      startColumnIndex,
      operations,
    });
  }
  private _assertColumnChangesAlone(): void {
    const otherWrites = {
      ...this.tableState.writeQueue,
      table: { ...this.writes, columnChanges: [] },
    };
    if (!hasQueuedWrites(otherWrites)) return;
    throw new Error(
      `Refusing to change the columns of ${this.tableLabel} alongside its other queued writes: they are indexed by the columns before the change. Flush the column change on its own.`,
    );
  }
  // Each change is read in the layout the ones before it leave, so the working view moves with it.
  private _columnChangeOperation(
    change: ColumnCountChange,
  ): TableColumnChangeOperation {
    const { colIndex } = change;
    const { kind } = columnChangeRequests[change.kind];
    const band = this._headAndBodyBand(colIndex, colIndex + 1);
    this._validateNoNeighbourSplitBy(
      band,
      this._columnChangeLabel(change.kind),
    );
    const shiftCount = columnCountChange[kind];
    this._shiftWorkingColumns(colIndex, shiftCount);
    this.growColumnCount(shiftCount);
    if (change.kind === "insertColumn") {
      this.column(colIndex)._initHeadCells(change.headCells);
    }
    return { kind, range: band, shiftDimension: "COLUMNS" };
  }
  // Not moveDimension, which reorders a Table stacked in the same sheet columns (sheets-behaviour.md).
  private _columnMoveOperations(
    move: ColumnMove,
  ): TableColumnChangeOperation[] {
    const { colIndex, count, beforeColIndex } = move;
    const gap = this._headAndBodyBand(beforeColIndex, beforeColIndex + count);
    this._validateNoNeighbourSplitBy(
      gap,
      this._columnChangeLabel("moveColumns"),
    );
    // An insert running past the grid's edge is clipped to it (sheets-behaviour.md), losing a cut column.
    this.sheet.queueGridColumnsThrough(gap.endColumnIndex);
    const colIndexAfterGap =
      colIndex > beforeColIndex ? colIndex + count : colIndex;
    const vacated = this._headAndBodyBand(
      colIndexAfterGap,
      colIndexAfterGap + count,
    );
    this._rekeyWorkingColumns((byColIndex) => movedColumns(byColIndex, move));
    return [
      { kind: "insertRange", range: gap, shiftDimension: "COLUMNS" },
      {
        kind: "cutPaste",
        source: vacated,
        destination: {
          sheetId: gap.sheetId,
          rowIndex: gap.startRowIndex,
          columnIndex: gap.startColumnIndex,
        },
      },
      { kind: "deleteRange", range: vacated, shiftDimension: "COLUMNS" },
    ];
  }
  // From the top head row down, so the head rows move with the Table.
  private _headAndBodyBand(
    startColIndex: number,
    endColIndex: number,
  ): BoundedGridRange {
    const origin = this.originAtGathering();
    return {
      sheetId: this.sheetGid,
      startRowIndex: origin.topHeadSheetRowIndex,
      endRowIndex: origin.sheetRowIndex(this.dataRowCount),
      startColumnIndex: origin.sheetColIndex(startColIndex),
      endColumnIndex: origin.sheetColIndex(endColIndex),
    };
  }
  private _shiftWorkingColumns(colIndex: number, shiftCount: 1 | -1): void {
    this._rekeyWorkingColumns((byColIndex) =>
      shiftedColumns(byColIndex, colIndex, shiftCount),
    );
    if (shiftCount !== 1) return;
    for (const rowState of this.rowStates.values()) {
      rowState.set(colIndex, { value: "" });
    }
  }
  private _rekeyWorkingColumns(
    rekey: <VL>(byColIndex: Map<number, VL>) => Map<number, VL>,
  ): void {
    for (const [rowIndex, rowState] of this.rowStates) {
      this.rowStates.set(rowIndex, rekey(rowState));
    }
    this.tableState.working.columnStates = rekey(this.columnStates);
    // An earlier change's new head cells are queued by colIndex, so a later change in the same gather carries them.
    for (const rowWrites of this.tableState.writeQueue.rows.values()) {
      rowWrites.fillCells = rekey(rowWrites.fillCells);
    }
  }
  private get _namedTableOnSheet(): string {
    return `Table "${this.name}" on sheet "${this.sheetTitle}"`;
  }
  private _columnChangeLabel(changeKind: TableColumnChange["kind"]): string {
    return `${columnChangeRequests[changeKind].action} ${this._namedTableOnSheet}`;
  }
  // Before the Table-end column inserts are counted, since growth is sent ahead of them.
  gatherAppendTableRowsOperation(): void {
    const appendedRowIndexes = this._queuedRowAppendIndexes();
    if (appendedRowIndexes.length === 0) return;
    const appendedRowCount = appendedRowIndexes.length;
    // Not dataRowCount: a same-run re-fetch resets the Table's end but keeps the queued appends.
    const modelRow = this.dataRowGridRange(Math.min(...appendedRowIndexes) - 1);
    const newRows: BoundedGridRange = {
      ...modelRow,
      startRowIndex: modelRow.endRowIndex,
      endRowIndex: SheetIndex.row(modelRow.endRowIndex + appendedRowCount),
    };
    // The appendDimension sent first already grows a Table ending on the grid's last row; an insert would push that row out of it.
    const endsOnGridEdge =
      newRows.startRowIndex === this.sheet.rowCountBeforeFlush;
    this.sheet.queueGridRowsThrough(newRows.endRowIndex);
    const insert: InsertRangeOperation[] = endsOnGridEdge
      ? []
      : [{ kind: "insertRange", range: newRows, shiftDimension: "ROWS" }];
    this.writeOperations.appendTableRows.push({
      tableId: this.tableId,
      newRows,
      operations: [
        ...insert,
        {
          kind: "updateTableRange",
          tableId: this.tableId,
          range: { ...newRows, startRowIndex: this.startRowIndex },
        },
        {
          kind: "copyPaste",
          source: modelRow,
          destination: newRows,
          pasteType: "PASTE_FORMAT",
        },
        ...columnRunCopies(
          modelRow,
          newRows,
          this._untypedColumnRuns(),
          "PASTE_DATA_VALIDATION",
        ),
        ...columnRunCopies(
          modelRow,
          newRows,
          this._appendedFormulaColumnRuns(),
          "PASTE_FORMULA",
        ),
      ],
    });
  }
  // An appended row deleted this flush still grows, then goes with the deletes, so the rows after it keep their place.
  private _queuedRowAppendIndexes(): number[] {
    return Array.from(this.rowWrites).flatMap(([rowIndex, writes]) =>
      writes.appendRow ? [rowIndex] : [],
    );
  }
  // A typed column's validation is the Table's own, so only Automatic columns get theirs copied.
  private _untypedColumnRuns(): ColumnRun[] {
    const columnTypes = this._columnTypesAfterFlush();
    return columnRuns(
      this.fullTableColIndexes.filter(
        (colIndex) => columnTypes.get(colIndex) === undefined,
      ),
    );
  }
  // Only the named columns, so the paste never reaches a plain value the append left unwritten.
  private _appendedFormulaColumnRuns(): ColumnRun[] {
    const { appendedFormulaColIndexes } = this.writes;
    return columnRuns(
      this.fullTableColIndexes.filter((colIndex) =>
        appendedFormulaColIndexes.has(colIndex),
      ),
    );
  }
  // A column type queued this flush is sent before growth, so it is the one the new rows meet.
  private _columnTypesAfterFlush(): Map<number, string | undefined> {
    const columnTypes = new Map(
      this.columnProperties.map(({ columnIndex, columnType }) => [
        columnIndex,
        columnType,
      ]),
    );
    const unfetched = this.fullTableColIndexes.find(
      (colIndex) => !columnTypes.has(colIndex),
    );
    if (unfetched !== undefined) {
      throw new Error(
        `${this.tableLabel} has no fetched type for ${tableColumnLabel(unfetched)}; refetch it before appending a row.`,
      );
    }
    this.writes.columnTypes.forEach((columnType, colIndex) =>
      columnTypes.set(colIndex, queuedColumnType.typeOf(columnType)),
    );
    return columnTypes;
  }
  // Google resets every dropdown's style and colours on a columnProperties send; observed: docs/sheets-behaviour.md#creating-and-updating-tables.
  // Counted as sent, so a dropdown the send removes, which has no style to keep, isn't.
  assertDropdownStyleResetAllowed(): void {
    const { columnTypes } = this.writes;
    if (columnTypes.size === 0) return;
    if (!this.spreadsheetStateRaw.rawSource.refusesDropdownStyleReset) return;
    const dropdowns = this._columnTypesColumnProperties(columnTypes).filter(
      ({ dataValidationRule }) => dataValidationRule !== undefined,
    );
    if (dropdowns.length === 0) return;
    throw new Error(
      `Refusing to set column types on ${this._tableLabel()} under dropdownStyleReset: "ask": it would reset the style and colours of its dropdowns ${dropdowns.map(({ columnName }) => columnName).join(", ")}.`,
    );
  }
  gatherSetTableColumnPropertiesOperation(): void {
    const { columnTypes } = this.writes;
    if (columnTypes.size === 0) return;
    this._assertColumnTypesUpdateAllowed(columnTypes);
    this.writeOperations.setTableColumnProperties.push({
      kind: "setTableColumnProperties",
      tableId: this.tableId,
      columnProperties: this._columnTypesColumnProperties(columnTypes),
    });
  }
  private _assertColumnTypesUpdateAllowed(columnTypes: ColumnTypes): void {
    const tableLabel = this._tableLabel();
    const fetched = this.columnProperties;
    if (fetched.length === 0) {
      throw new Error(
        `${tableLabel} has no fetched column properties; refetch it before setting a column type.`,
      );
    }
    if (
      this.writeOperations.insertTableEndColumns.some(
        ({ newColumns }) => newColumns.sheetId === this.sheetGid,
      ) ||
      this.writeOperations.changeTableColumns.some(
        ({ sheetId }) => sheetId === this.sheetGid,
      )
    ) {
      throw new Error(
        `Refusing to set column types on ${tableLabel}: the same flush inserts or deletes a column on that sheet.`,
      );
    }
    columnTypes.forEach((_, columnIndex) => {
      if (!fetched.some((column) => column.columnIndex === columnIndex)) {
        throw new Error(
          `${tableLabel} has no fetched ${tableColumnLabel(columnIndex)}.`,
        );
      }
    });
  }
  // Full list with each rule, since a partial columnProperties replaces the rest and a dropdown sent without its rule is refused.
  private _columnTypesColumnProperties(
    columnTypes: ColumnTypes,
  ): TableColumnPropertiesUpdate[] {
    return this.columnProperties.map((column) => {
      const { columnIndex } = column;
      if (column.columnName === undefined) {
        throw new Error(
          `${this._tableLabel()} ${tableColumnLabel(columnIndex)} has no columnName; refusing to replace column properties.`,
        );
      }
      const queued = columnTypes.get(columnIndex);
      const baseline =
        queued === undefined ? undefined : queuedColumnType.baselineOf(queued);
      if (baseline !== undefined) {
        return {
          columnIndex,
          columnName: column.columnName,
          ...queuedColumnType.properties(baseline),
        };
      }
      const columnType =
        queued === undefined
          ? column.columnType
          : queuedColumnType.typeOf(queued);
      const dataValidationRule = validationRuleOf(column);
      return {
        columnIndex,
        columnName: column.columnName,
        ...(columnType !== undefined ? { columnType } : {}),
        ...(dataValidationRule !== undefined ? { dataValidationRule } : {}),
      };
    });
  }
  private _tableLabel(): string {
    return `Table ${this.tableId} on "${this.sheet.title}"`;
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
  private _sheetConditionalFormatRule(
    { range, condition, format }: TableConditionalFormatRule,
    origin: TableOrigin,
  ): ModelableConditionalFormatRule {
    return {
      kind: "boolean",
      ranges: [this._sheetGridRange(range, origin)],
      condition,
      format,
    };
  }
  private _sheetEditProtection(
    { range, ...protection }: TableEditProtection,
    origin: TableOrigin,
  ): EditProtectionContent {
    return { ...protection, range: this._sheetGridRange(range, origin) };
  }
  // Absent bounds stay absent, as the matching of fetched rules and protections expects.
  private _sheetGridRange(
    { startRowIndex, endRowIndex, startColIndex, endColIndex }: TableGridRange,
    origin: TableOrigin,
  ): GridRangeProps {
    const range: GridRangeProps = {
      sheetId: this.sheetGid,
      startRowIndex: SheetIndex.row(0),
    };
    if (startRowIndex !== undefined) {
      range.startRowIndex = origin.sheetRowIndex(startRowIndex);
    }
    if (endRowIndex !== undefined) {
      range.endRowIndex = origin.sheetRowIndex(endRowIndex);
    }
    if (startColIndex !== undefined) {
      range.startColumnIndex = origin.sheetColIndex(startColIndex);
    }
    if (endColIndex !== undefined) {
      range.endColumnIndex = origin.sheetColIndex(endColIndex);
    }
    return range;
  }
  private _queueEditProtection(protection: TableEditProtection): this {
    // A same-batch shift carries only a Table whose properties have arrived.
    this.assertTableIsKnown();
    if (hasTableRows(protection.range)) this.sheet.assertRowIndexesNotStale();
    this.sheet.assertEditProtectionsNotStale();
    const queued = this.tablesOnSheet.flatMap((table) =>
      table._queuedSheetEditProtections(),
    );
    if (
      this.sheet.hasPendingEditProtection(
        this._sheetEditProtection(protection, this.tableOrigin()),
        queued,
      )
    ) {
      return this;
    }
    return this.queueTableWrite({ action: "addEditProtection", ...protection });
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
  private _queuedSheetConditionalFormatRules(): ModelableConditionalFormatRule[] {
    const origin = this.tableOrigin();
    return this.writes.conditionalFormatRules.map((rule) =>
      this._sheetConditionalFormatRule(rule, origin),
    );
  }
  private _queuedSheetEditProtections(): EditProtectionContent[] {
    const origin = this.tableOrigin();
    return this.writes.editProtections.map((protection) =>
      this._sheetEditProtection(protection, origin),
    );
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

type ColumnTypes = TableWrites["columnTypes"];

// A whole column has none, so moved rows don't bar it.
function hasTableRows(range: TableGridRange): boolean {
  return range.startRowIndex !== undefined;
}

function splittingPhrase(actionLabel: string, neighbourName: string): string {
  return `${actionLabel} would shift only part of Table "${neighbourName}"`;
}

function shiftedColumns<VL>(
  byColIndex: Map<number, VL>,
  colIndex: number,
  shiftCount: 1 | -1,
): Map<number, VL> {
  return remappedColumns(byColIndex, (index) => {
    if (index < colIndex) return index;
    if (index > colIndex || shiftCount === 1) return index + shiftCount;
    return undefined;
  });
}

function movedColumns<VL>(
  byColIndex: Map<number, VL>,
  span: ColumnMoveSpan,
): Map<number, VL> {
  return remappedColumns(byColIndex, (index) => movedColIndex(index, span));
}

// A column the remap gives no index goes.
function remappedColumns<VL>(
  byColIndex: Map<number, VL>,
  newColIndex: (index: number) => number | undefined,
): Map<number, VL> {
  return Array.from(byColIndex).reduce((remapped, [index, value]) => {
    const remappedIndex = newColIndex(index);
    if (remappedIndex !== undefined) remapped.set(remappedIndex, value);
    return remapped;
  }, new Map<number, VL>());
}

// Where a column sits once the block has moved; the columns it passes close the gap.
function movedColIndex(
  index: number,
  { colIndex, count, beforeColIndex }: ColumnMoveSpan,
): number {
  const isInBlock = colIndex <= index && index < colIndex + count;
  if (beforeColIndex > colIndex) {
    if (isInBlock) return index + beforeColIndex - colIndex - count;
    if (colIndex + count <= index && index < beforeColIndex) {
      return index - count;
    }
    return index;
  }
  if (isInBlock) return index - (colIndex - beforeColIndex);
  if (beforeColIndex <= index && index < colIndex) return index + count;
  return index;
}

type ColumnRun = [startColIndex: number, endColIndex: number];

// Ascending indexes, so each run of neighbours becomes one range.
function columnRuns(colIndexes: number[]): ColumnRun[] {
  return colIndexes.reduce<ColumnRun[]>((runs, colIndex) => {
    const lastRun = runs.at(-1);
    if (lastRun !== undefined && lastRun[1] === colIndex) {
      lastRun[1] = colIndex + 1;
    } else {
      runs.push([colIndex, colIndex + 1]);
    }
    return runs;
  }, []);
}

function columnRunCopies(
  source: BoundedGridRange,
  destination: BoundedGridRange,
  runs: ColumnRun[],
  pasteType: CopyPasteType,
): CopyPasteOperation[] {
  return runs.map(([startColIndex, endColIndex]) => ({
    kind: "copyPaste",
    source: columnRun(source, startColIndex, endColIndex),
    destination: columnRun(destination, startColIndex, endColIndex),
    pasteType,
  }));
}

function columnRun(
  rows: BoundedGridRange,
  startColIndex: number,
  endColIndex: number,
): BoundedGridRange {
  return {
    ...rows,
    startColumnIndex: SheetIndex.col(rows.startColumnIndex + startColIndex),
    endColumnIndex: SheetIndex.col(rows.startColumnIndex + endColIndex),
  };
}

// Google reads every rule back with its condition type, so a column with none has no rule to resend.
function validationRuleOf(
  column: TableColumnSnapshot,
): TableColumnValidationRule | undefined {
  const { dataValidationConditionType, dataValidationValues } = column;
  if (dataValidationConditionType === undefined) return undefined;
  return {
    condition: dataValidationConditionType,
    values: dataValidationValues,
  };
}

function tableColumnLabel(columnIndex: number): string {
  return `table column ${columnIndex}`;
}
