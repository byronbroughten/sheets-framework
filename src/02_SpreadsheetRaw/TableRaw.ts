import type { CellValueName } from "../00_Source/CellValues/cellValues";
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
  GridBlockSnapshot,
  GridCellSnapshot,
  GridRangeProps,
  SheetSnapshot,
  TableColumnPropertiesUpdate,
  TableColumnSnapshot,
} from "../00_Source/RawSource/RawSource";
import { SheetIndex } from "../00_Source/RawSource/SheetIndex";
import { type HeadRole, headRows } from "../01_SpreadsheetSchema/headRows";
import type { TableOrigin } from "../01_SpreadsheetSchema/TableOrigin";
import type { Value } from "../01_SpreadsheetSchema/valueSchemas";
import { Arr } from "../utils/Arr";
import { Val } from "../utils/Val";
import { assertValueAndFormulaExclusive } from "./CellRaw";
import type { RowCommonRaw } from "./ClassBases/RowCommonRaw";
import { originOf } from "./ClassBases/TableBaseRaw";
import { TableCommonRaw } from "./ClassBases/TableCommonRaw";
import {
  type ColumnFill,
  type FindReplaceTerms,
  type SortParameters,
  type TableCell,
  type TableConditionalFormatRule,
  type TableEditProtection,
  type TableEndColumnHeadCells,
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

/**
 * One Table's state by Table-relative index: rows, columns, pruning, queued
 * Table-level requests, and integrating fetched cells into its rows, cells and
 * sampled column facts. Its sheet's title, grid size, conditional format rules
 * and edit protections are SheetRaw, reached as `table.sheet`.
 * Descriptive facts are TableProfileRaw and ColumnProfileRaw; column ID lookups are TableRaw/;
 * spreadsheet-wide fetch and flush are SpreadsheetRaw. By-name and columnId resolution are Identified/Named.
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
    const origin = this.originAtGathering();
    const { columnCount } = this;
    return {
      sheetId: this.sheetGid,
      startRowIndex: origin.sheetRowIndex(rowIndex),
      endRowIndex: origin.sheetRowIndex(rowIndex + 1),
      startColumnIndex: origin.sheetColIndex(0),
      endColumnIndex: origin.sheetColIndex(columnCount),
    };
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
  appendColumn(headCells: TableEndColumnHeadCells): number {
    const colIndex = this.columnCount + this.writes.insertTableEndColumnCount;
    this.queueTableWrite({ action: "insertTableEndColumn" });
    this.column(colIndex)._initHeadCells(headCells);
    return colIndex;
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
    // The live start is unknown until this probe comes back, so it aims where the configs record the Table.
    const origin = this.presumedOrigin;
    this.gatherFetchRange({
      startRowIndex: SheetIndex.row(0),
      endRowIndex: SheetIndex.row(origin.headerRowIndex + 1),
      startColumnIndex: origin.startColIndex,
      endColumnIndex: SheetIndex.col(origin.startColIndex + 1),
    });
    headRows.indexes().forEach((rowIndex) => {
      this.headRowByIndex(rowIndex).cell(0).prepFetchBackfill();
    });
    const { recordedTableId } = this;
    if (recordedTableId !== undefined) {
      this.sheetState.fetchQueue.placementStripTableIds.add(recordedTableId);
    }
    return this;
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
  integrateSheetState(sheet: SheetSnapshot): void {
    this._integrateSheetProperties(sheet);
    this.tableIds().forEach((tableId) => {
      this.ss.table(tableId).integrateGridBlocks(sheet.gridBlocks ?? []);
    });
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
    const queued = this._tablesOnSheet().flatMap((table) =>
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
  requestSortGSheet({ colIdxToSortBy, sortOrder }: SortParameters): void {
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
  }
  gatherCheckboxValidationOperation({ rowIndex, colIndex }: TableCell): void {
    const origin = this.originAtGathering();
    this.writeOperations.addCheckboxValidation.push({
      kind: "addCheckboxValidation",
      range: {
        sheetId: this.sheetGid,
        startRowIndex: origin.sheetRowIndex(rowIndex),
        endRowIndex: origin.sheetRowIndex(rowIndex + 1),
        startColumnIndex: origin.sheetColIndex(colIndex),
        endColumnIndex: origin.sheetColIndex(colIndex + 1),
      },
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
    return `Inserting a column at the end of Table "${this.name}" on sheet "${this.sheetTitle}" would shift only part of Table "${neighbourName}"`;
  }
  // From the top head row down, so the head rows move with the Table; the widen is what grows it.
  gatherInsertTableEndColumnsOperation(): void {
    const insertCount = this.writes.insertTableEndColumnCount;
    if (insertCount === 0) return;
    const origin = this.originAtGathering();
    const { columnCount, dataRowCount } = this;
    const newColumns: BoundedGridRange = {
      sheetId: this.sheetGid,
      startRowIndex: origin.topHeadSheetRowIndex,
      endRowIndex: origin.sheetRowIndex(dataRowCount),
      startColumnIndex: origin.sheetColIndex(columnCount),
      endColumnIndex: origin.sheetColIndex(columnCount + insertCount),
    };
    this._validateNoNeighbourSplitBy(newColumns);
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
  private _validateNoNeighbourSplitBy(band: BoundedGridRange): void {
    const split = this.tableIds()
      .filter((tableId) => tableId !== this.tableId)
      .map((tableId) => this.ss.table(tableId))
      .find((neighbour) => neighbour.isSplitBy(band));
    if (split === undefined) return;
    throw new Error(
      `${this.columnInsertSplitting(split.name)} with its head rows. Move "${split.name}" so that it and its head rows sit entirely within rows ${band.startRowIndex + 1}–${band.endRowIndex}, or entirely outside them.`,
    );
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
    this.sheet.queueGridRowsThrough(newRows.endRowIndex);
    this.writeOperations.appendTableRows.push({
      tableId: this.tableId,
      newRows,
      operations: [
        { kind: "insertRange", range: newRows, shiftDimension: "ROWS" },
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
        ...this._untypedColumnRuns().map(
          ([startColIndex, endColIndex]): CopyPasteOperation => ({
            kind: "copyPaste",
            source: columnRun(modelRow, startColIndex, endColIndex),
            destination: columnRun(newRows, startColIndex, endColIndex),
            pasteType: "PASTE_DATA_VALIDATION",
          }),
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
  private _untypedColumnRuns(): [number, number][] {
    const columnTypes = this._columnTypesAfterFlush();
    return this.fullTableColIndexes.reduce<[number, number][]>(
      (runs, colIndex) => {
        if (columnTypes.get(colIndex) !== undefined) return runs;
        const lastRun = runs.at(-1);
        if (lastRun !== undefined && lastRun[1] === colIndex) {
          lastRun[1] = colIndex + 1;
        } else {
          runs.push([colIndex, colIndex + 1]);
        }
        return runs;
      },
      [],
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
      columnTypes.set(colIndex, columnType),
    );
    return columnTypes;
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
    const tableLabel = this._tableLabel(this.tableId);
    const fetched = this.columnProperties;
    if (fetched.length === 0) {
      throw new Error(
        `${tableLabel} has no fetched column properties; refetch it before setting a column type.`,
      );
    }
    if (
      this.writeOperations.insertTableEndColumns.some(
        ({ newColumns }) => newColumns.sheetId === this.sheetGid,
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
    columnTypes.forEach((_, columnIndex) => {
      if (!fetched.some((column) => column.columnIndex === columnIndex)) {
        throw new Error(
          `${tableLabel} has no fetched ${tableColumnLabel(columnIndex)}.`,
        );
      }
    });
  }
  // Full list, since a partial columnProperties replaces the rest.
  private _columnTypesColumnProperties(
    columnTypes: ColumnTypes,
  ): TableColumnPropertiesUpdate[] {
    return this.columnProperties.map((column) => {
      const { columnIndex } = column;
      if (column.columnName === undefined) {
        throw new Error(
          `${this._tableLabel(this.tableId)} ${tableColumnLabel(columnIndex)} has no columnName; refusing to replace column properties.`,
        );
      }
      const columnType = columnTypes.get(columnIndex) ?? column.columnType;
      return {
        columnIndex,
        columnName: column.columnName,
        ...(columnType !== undefined ? { columnType } : {}),
      };
    });
  }
  private _tableLabel(tableId: string): string {
    return `Table ${tableId} on "${this.sheet.title}"`;
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
    const origin = this.originAtGathering();
    return {
      sheetId: this.sheetGid,
      startRowIndex: origin.sheetRowIndex(0),
      endRowIndex: origin.sheetRowIndex(dataRowCount),
      startColumnIndex: origin.sheetColIndex(0),
      endColumnIndex: origin.sheetColIndex(this.columnCount),
    };
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
    if (hasTableRows(protection.range)) this.sheet.assertRowIndexesNotStale();
    this.sheet.assertEditProtectionsNotStale();
    const queued = this._tablesOnSheet().flatMap((table) =>
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
  // This Table even before it is fetched, plus every known neighbour, since a head row's range spans them all.
  private _tablesOnSheet(): TableRaw[] {
    const neighbours = this.tableIds()
      .map((tableId) => this.ss.table(tableId))
      .filter(
        (table) => !this.hasFetchedProperties || table.tableId !== this.tableId,
      );
    return [this, ...neighbours];
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
  appendDataRow(): RowRaw {
    return this.row(this.dataRowCount).append();
  }
  appendDataRowValues(colValues: Map<number, Value>): RowRaw {
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

function columnLabel(column: TableColumnSnapshot): string {
  return column.columnName ?? tableColumnLabel(column.columnIndex);
}

function tableColumnLabel(columnIndex: number): string {
  return `table column ${columnIndex}`;
}
