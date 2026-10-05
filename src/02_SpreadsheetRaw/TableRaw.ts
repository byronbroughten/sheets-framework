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
  CopyPasteOperation,
  GridBlockSnapshot,
  GridRangeProps,
  SheetSnapshot,
  TableColumnPropertiesUpdate,
  TableColumnSnapshot,
} from "../00_Source/RawSource/RawSource";
import {
  SheetIndex,
  type SheetRowIndex,
} from "../00_Source/RawSource/SheetIndex";
import { TableOrigin } from "../01_SpreadsheetSchema/TableOrigin";
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
  type TableWrites,
} from "./ClassTypes/StateRaw";
import { ColumnRaw } from "./ColumnRaw";
import { RowRaw } from "./RowRaw";
import { SheetMetaRaw } from "./SheetMetaRaw";
import { SpreadsheetRaw } from "./SpreadsheetRaw";
import { SheetConditionalFormatsRaw } from "./TableRaw/SheetConditionalFormatsRaw";
import { SheetEditProtectionsRaw } from "./TableRaw/SheetEditProtectionsRaw";

/**
 * One Table's state by Table-relative index: rows, columns, pruning, queued
 * Table-level requests, and integrating fetched cells into its rows, cells and
 * Meta column facts. `ss.sheet(gid)` also reaches it through its sheet, so
 * sheet-level title, conditional format rules and edit protections live here
 * too, the latter two in TableRaw/ behind one-line delegations.
 * Uniform rows and column facts are SheetMetaRaw; spreadsheet-wide fetch and
 * flush are SpreadsheetRaw. By-name and columnId resolution are Identified/Named.
 */
export class TableRaw extends TableCommonRaw {
  get ss(): SpreadsheetRaw {
    return new SpreadsheetRaw(this.spreadsheetRawProps);
  }
  get meta(): SheetMetaRaw {
    return new SheetMetaRaw(this.tableRawProps);
  }
  private get conditionalFormats(): SheetConditionalFormatsRaw {
    return new SheetConditionalFormatsRaw(this.tableRawProps);
  }
  private get protections(): SheetEditProtectionsRaw {
    return new SheetEditProtectionsRaw(this.tableRawProps);
  }
  get hasFetchedProperties(): boolean {
    return this.tableProperties !== undefined;
  }
  dataGridRange(): GridRangeProps {
    const { origin, dataRowCount, columnCount } = this;
    return {
      sheetId: this.sheetGid,
      startRowIndex: origin.sheetRowIndex(0),
      endRowIndex: origin.sheetRowIndex(dataRowCount),
      startColumnIndex: origin.sheetColIndex(0),
      endColumnIndex: origin.sheetColIndex(columnCount),
    };
  }
  get wholeSheetGridRange(): ProtectionGridRange {
    return { sheetId: this.sheetGid };
  }
  rowGridRange(rowIndex: number): GridRangeProps {
    const origin = this.tableOrigin();
    return {
      sheetId: this.sheetGid,
      startRowIndex: origin.sheetRowIndex(rowIndex),
      endRowIndex: origin.sheetRowIndex(rowIndex + 1),
    };
  }
  // The live Table's columns only, so a range built on it never reaches a neighbour.
  dataRowGridRange(rowIndex: number): BoundedGridRange {
    const { origin, columnCount } = this;
    return {
      sheetId: this.sheetGid,
      startRowIndex: origin.sheetRowIndex(rowIndex),
      endRowIndex: origin.sheetRowIndex(rowIndex + 1),
      startColumnIndex: origin.sheetColIndex(0),
      endColumnIndex: origin.sheetColIndex(columnCount),
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
    this.writeOperations.renameSheet.push({
      kind: "renameSheet",
      sheetId: this.sheetGid,
      title,
    });
    this.sheetState.working.title = title;
    return this;
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
  get activeRowIndexes(): number[] {
    const indexes = Array.from(this.rowStates.keys());
    return Arr.sortAscending(indexes);
  }
  get activeRowCount(): number {
    return this.rowStates.size;
  }
  get lastActiveRowIndex(): number {
    return Math.max(...this.rowStates.keys());
  }
  get rowIndexesActive(): number[] {
    return this.activeRowIndexes.filter((rowIndex) => rowIndex >= 0);
  }
  get rowIndexesFull(): number[] {
    return Arr.indexesFromUntil(0, this.dataRowCount);
  }
  get rowsFull(): RowRaw[] {
    return this.rowIndexesFull.map((rowIndex) => this.row(rowIndex));
  }
  get rows(): RowRaw[] {
    return this.rowIndexesActive.map((index) => this.row(index));
  }
  get topRow(): RowRaw {
    return this.row(0);
  }
  get rowCount(): number {
    return this.rowIndexesActive.length;
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
  findReplace(terms: FindReplaceTerms): this {
    this.ss.findReplace({ ...terms, scope: { sheetId: this.sheetGid } });
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
      ...this.tableRawProps,
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
  gatherFetchProperties(): this {
    // The live start is unknown until this probe comes back, so it aims where the layout expects the Table.
    const origin = TableOrigin.expected();
    this.gatherFetchRange({
      startRowIndex: SheetIndex.row(0),
      endRowIndex: SheetIndex.row(origin.headerRowIndex + 1),
      startColumnIndex: origin.startColIndex,
      endColumnIndex: SheetIndex.col(origin.startColIndex + 1),
    });
    this.schema.uniformRowNames.forEach((name) => {
      this.meta.uniformRow(name).cell(0).prepFetchBackfill();
    });
    this.sheetState.fetchQueue.gatherPlacementStrip = true;
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
      this.rowCommon(rowIndex).ensureFullActiveDataCells();
    });
    toFinalize.columns.forEach((colIndex) => {
      this.column(colIndex).ensureFullActiveDataCells();
    });
    this._ensureFetchedActiveFacts();
    toFinalize.rows.clear();
    toFinalize.columns.clear();
  }
  private _finalizeFetchedCells(): void {
    this.tableState.fetchQueue.toFinalize.cells.forEach(
      (colIndexes, rowIndex) => {
        const row = this.rowCommon(rowIndex);
        row.ensureStateExists();
        colIndexes.forEach((colIndex) => {
          row.cell(colIndex).ensureActive();
        });
      },
    );
    this.tableState.fetchQueue.toFinalize.cells.clear();
  }
  // After the backfills above, so a blank fact is sampled rather than built.
  private _ensureFetchedActiveFacts(): void {
    const { toFinalize } = this.tableState.fetchQueue;
    if (toFinalize.rows.has(0)) {
      this.meta.ensureTableColumnsActiveFacts();
    }
    toFinalize.columns.forEach((colIndex) => {
      if (!this.isTableColIndex(colIndex)) return;
      this.meta.column(colIndex).ensureActiveFacts();
    });
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
          if (row.rowIsActive()) {
            row.cell(colIndex).integrateSnapshot(cellData);
          }
          if (rowIndex === 0) {
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
    this.tableState.working.isPrunedToSelection = true;
  }
  // A whole-column fill ignores active rows, so it would rewrite what a prune excluded.
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
    const origin = this.tableOrigin();
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
  addCheckboxValidationAt(range: BoundedGridRange): this {
    this.assertRowIndexesNotStale();
    this.writeOperations.addCheckboxValidation.push({
      kind: "addCheckboxValidation",
      range,
    });
    return this;
  }
  gatherQueuedTableWrites(): void {
    const { writes } = this;
    this.gatherInsertTableEndColumnOperations(writes.insertTableEndColumnCount);
    if (writes.sort !== undefined) {
      this.gatherSortOperation(writes.sort);
    }
    writes.fillColumns.forEach((fill) => {
      this.gatherFillColumnOperation(fill);
    });
  }
  gatherInsertTableEndColumnOperations(insertCount: number): void {
    Array.from({ length: insertCount }).forEach(() => {
      const { origin, columnCount } = this;
      this.writeOperations.insertTableEndColumn.push({
        kind: "insertTableEndColumn",
        sheetId: this.sheetGid,
        startColumnIndex: origin.sheetColIndex(columnCount),
      });
      this.growColumnCount();
    });
  }
  // Before the Table-end column inserts are counted, since growth is sent ahead of them.
  gatherAppendTableRowsOperation(): void {
    const appendedRowIndexes = this._queuedRowAppendIndexes();
    if (appendedRowIndexes.length === 0) return;
    this.assertRowIndexesNotStale();
    const appendedRowCount = appendedRowIndexes.length;
    // Not dataRowCount: a same-run re-fetch resets the Table's end but keeps the queued appends.
    const modelRow = this.dataRowGridRange(Math.min(...appendedRowIndexes) - 1);
    const newRows: BoundedGridRange = {
      ...modelRow,
      startRowIndex: modelRow.endRowIndex,
      endRowIndex: SheetIndex.row(modelRow.endRowIndex + appendedRowCount),
    };
    this._queueGridRowsThrough(newRows.endRowIndex);
    this.writeOperations.appendTableRows.push({
      sheetId: this.sheetGid,
      startRowIndex: newRows.startRowIndex,
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
  private _queuedRowAppendIndexes(): number[] {
    return Array.from(this.rowWrites).flatMap(([rowIndex, writes]) =>
      writes.appendRow && !writes.deleteRow ? [rowIndex] : [],
    );
  }
  private _queueGridRowsThrough(endRowIndex: SheetRowIndex): void {
    const { working, writeQueue } = this.sheetState;
    const rowCount = Val.assert(
      working.rowCount,
      `${this.sheetLabel}'s row count`,
    );
    if (endRowIndex <= rowCount) return;
    writeQueue.appendedRowCount += endRowIndex - rowCount;
    working.rowCount = endRowIndex;
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
      this.writeOperations.insertTableEndColumn.some(
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
    return `Table ${tableId} on "${this.title}"`;
  }
  gatherSortOperation({ colIdxToSortBy, sortOrder }: SortParameters): void {
    const origin = this.tableOrigin();
    this.writeOperations.sort.push({
      kind: "sort",
      sheetId: this.sheetGid,
      startRowIndex: origin.sheetRowIndex(0),
      startColumnIndex: origin.sheetColIndex(0),
      colIdxToSortBy: origin.sheetColIndex(colIdxToSortBy),
      sortOrder,
    });
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
