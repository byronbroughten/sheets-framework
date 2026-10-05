import type {
  BoundedGridRange,
  TableColumnSnapshot,
} from "../../00_Source/RawSource/RawSource";
import {
  type SheetColIndex,
  SheetIndex,
  type SheetRowIndex,
} from "../../00_Source/RawSource/SheetIndex";
import type { TableOrigin } from "../../01_SpreadsheetSchema/TableOrigin";
import { Arr } from "../../utils/Arr";
import { Obj } from "../../utils/Obj";
import type { SheetGridRangeProps } from "../ClassTypes/AccessorsRaw";
import { emptyStateRaw } from "../ClassTypes/emptyStateRaw";
import type {
  AppendTableRows,
  CellFill,
  ColumnFill,
  TablePropertiesRaw,
  TableWriteProps,
  TableWriteQueueRaw,
  TableWrites,
} from "../ClassTypes/StateRaw";
import type { SpreadsheetRaw } from "../SpreadsheetRaw";
import { originOf, TableBaseRaw } from "./TableBaseRaw";

// Not on TableBaseRaw: the row and column base classes hang off that.
export abstract class TableCommonRaw extends TableBaseRaw {
  // Abstract: importing SpreadsheetRaw here would close an init-time cycle.
  abstract get ss(): SpreadsheetRaw;
  get tableId(): string {
    return this._knownTableProperties().tableId;
  }
  get name(): string {
    return this._knownTableProperties().name;
  }
  get origin(): TableOrigin {
    return originOf(this._workingTableProperties());
  }
  get startRowIndex(): SheetRowIndex {
    return this._workingTableProperties().startRowIndex;
  }
  get startColumnIndex(): SheetColIndex {
    return this._workingTableProperties().startColumnIndex;
  }
  get isHeaderOnly(): boolean {
    const { startRowIndex, endRowIndex } = this._knownTableProperties();
    return endRowIndex <= startRowIndex + 1;
  }
  get tableLabel(): string {
    return `Table "${this.name}" on ${this.sheetLabel}`;
  }
  get headerOnlyFix(): string {
    return `${this.tableLabel} has only its header: add a row below it holding its formulas.`;
  }
  get dataRowCount(): number {
    this.assertRowIndexesNotStale();
    const { startRowIndex, endRowIndex } = this._workingTableProperties();
    return endRowIndex - startRowIndex - 1;
  }
  get columnCount(): number {
    const { startColumnIndex, endColumnIndex } =
      this._workingTableProperties();
    return endColumnIndex - startColumnIndex;
  }
  get columnProperties(): TableColumnSnapshot[] {
    return this._workingTableProperties().columnProperties;
  }
  get rowIndexesAreStale(): boolean {
    return this.tableProperties?.rowIndexesAreStale === true;
  }
  get fullTableColIndexes(): number[] {
    return Arr.indexesFromUntil(0, this.columnCount);
  }
  get writes(): TableWrites {
    return this.tableState.writeQueue.table;
  }
  get rowWrites(): TableWriteQueueRaw["rows"] {
    return this.tableState.writeQueue.rows;
  }
  growDataRowCount(): void {
    this.assertRowIndexesNotStale();
    const properties = this._workingTableProperties();
    properties.endRowIndex = SheetIndex.row(properties.endRowIndex + 1);
  }
  growColumnCount(): void {
    const properties = this._workingTableProperties();
    properties.endColumnIndex = SheetIndex.col(properties.endColumnIndex + 1);
  }
  // An insert over any of this Table's columns, at or above its header, pushes it down.
  rowShiftFrom(growths: AppendTableRows[]): number {
    const properties = this.tableProperties;
    if (properties === undefined) return 0;
    return growths
      .filter(({ newRows }) => newRows.sheetId === this.sheetGid)
      .filter(({ newRows }) => isBelow(properties, newRows))
      .reduce(
        (rowCount, { newRows }) =>
          rowCount + newRows.endRowIndex - newRows.startRowIndex,
        0,
      );
  }
  shiftRowsDown(rowCount: number): void {
    const properties = this._knownTableProperties();
    properties.startRowIndex = SheetIndex.row(
      properties.startRowIndex + rowCount,
    );
    properties.endRowIndex = SheetIndex.row(properties.endRowIndex + rowCount);
  }
  markRowIndexesStale(): void {
    this._knownTableProperties().rowIndexesAreStale = true;
  }
  clearRowIndexStale(): void {
    this._workingTableProperties().rowIndexesAreStale = false;
  }
  // The sent list is now the Table's, and what was fetched no longer is.
  markColumnPropertiesStale(): void {
    this._workingTableProperties().columnProperties = [];
  }
  assertTableIsKnown(): void {
    this._workingTableProperties();
  }
  assertRowIndexesNotStale(): void {
    if (!this._workingTableProperties().rowIndexesAreStale) return;
    throw new Error(
      `Row indexes are stale for ${this.tableLabel}: a flush has moved its rows, so it needs a refetch, in a new run, before another row write.`,
    );
  }
  // The table's own range, not the layout's: no table means no table columns.
  isTableColIndex(colIndex: number): boolean {
    if (this.tableProperties === undefined) return false;
    return colIndex >= 0 && colIndex < this.columnCount;
  }
  gatherFetchRange(gr: SheetGridRangeProps): this {
    this.spreadsheetStateRaw.fetchQueue.gridRanges.push({
      sheetId: this.sheetGid,
      ...gr,
    });
    return this;
  }
  gatherFetchRanges(props: SheetGridRangeProps[]): this {
    props.forEach((props) => this.gatherFetchRange(props));
    return this;
  }
  // A full row is the Table's columns; before the Table is known, all columns from where the layout expects it.
  fullRowFetchRange(rowIndex: number): SheetGridRangeProps {
    const origin = this.tableOrigin();
    const range = {
      startRowIndex: origin.sheetRowIndex(rowIndex),
      endRowIndex: origin.sheetRowIndex(rowIndex + 1),
      startColumnIndex: origin.sheetColIndex(0),
    };
    const properties = this.tableProperties;
    if (properties === undefined) return range;
    return { ...range, endColumnIndex: properties.endColumnIndex };
  }
  // A full column runs from the column ID row, the topmost head row, to the Table's last row.
  fullColumnFetchRange(colIndex: number): SheetGridRangeProps {
    const origin = this.tableOrigin();
    const range = {
      startRowIndex: origin.sheetRowIndex(this.schema.colIdRowIndex),
      startColumnIndex: origin.sheetColIndex(colIndex),
      endColumnIndex: origin.sheetColIndex(colIndex + 1),
    };
    const properties = this.tableProperties;
    if (properties === undefined) return range;
    return { ...range, endRowIndex: properties.endRowIndex };
  }
  queueTableWrite(props: TableWriteProps): this {
    const writes = this.writes;
    switch (props.action) {
      case "sort":
        writes.sort = {
          colIdxToSortBy: props.colIdxToSortBy,
          sortOrder: props.sortOrder,
        };
        break;
      case "insertTableEndColumn":
        writes.insertTableEndColumnCount++;
        break;
      case "fillColumn": {
        const fill = Obj.strictOmit(props, "action");
        this._eraseCellFieldsUnder(fill);
        writes.fillColumns.push(fill);
        break;
      }
      case "updateColumnType":
        writes.columnTypes.set(props.colIndex, props.columnType);
        break;
      default:
        throw new Error(
          `Invalid action: ${(props as TableWriteProps).action}. Must be one of "sort", "insertTableEndColumn", "fillColumn" or "updateColumnType".`,
        );
    }
    return this;
  }
  // Column fills are sent before cell writes, so a later column fill wins by erasing what it covers.
  private _eraseCellFieldsUnder(fill: ColumnFill): void {
    for (const [rowIndex, rowWrites] of this.tableState.writeQueue.rows) {
      if (rowIndex < fill.startRowIndex || rowIndex >= fill.endRowIndex) {
        continue;
      }
      const cellFill = rowWrites.fillCells.get(fill.colIndex);
      if (cellFill === undefined) continue;
      const fieldsLeft = cellFieldsLeftUnder(fill, cellFill);
      if (Object.keys(fieldsLeft).length === 0) {
        rowWrites.fillCells.delete(fill.colIndex);
      } else {
        rowWrites.fillCells.set(fill.colIndex, fieldsLeft);
      }
    }
  }
  // The flusher's step once it has gathered this Table's queue into operations.
  _clearWriteQueue(): void {
    this.tableState.writeQueue.table = emptyStateRaw.tableWrites();
    this.tableState.writeQueue.rows = new Map();
  }
  private _knownTableProperties(): TablePropertiesRaw {
    const properties = this.tableProperties;
    if (properties === undefined) {
      throw new Error(
        `Table is unknown for sheetGid ${this.sheetGid}. Ensure that the sheet properties have been fetched.`,
      );
    }
    return properties;
  }
  // A Table met with only its header has no row to work in.
  private _workingTableProperties(): TablePropertiesRaw {
    if (this.isHeaderOnly) throw new Error(this.headerOnlyFix);
    return this._knownTableProperties();
  }
}

function isBelow(
  properties: TablePropertiesRaw,
  newRows: BoundedGridRange,
): boolean {
  return (
    properties.startRowIndex >= newRows.startRowIndex &&
    properties.startColumnIndex < newRows.endColumnIndex &&
    newRows.startColumnIndex < properties.endColumnIndex
  );
}

function cellFieldsLeftUnder(fill: ColumnFill, cellFill: CellFill): CellFill {
  const fieldsLeft = { ...cellFill };
  if (fill.value !== undefined || fill.formula !== undefined) {
    delete fieldsLeft.value;
    delete fieldsLeft.formula;
  }
  if (fill.backgroundColor !== undefined) delete fieldsLeft.backgroundColor;
  return fieldsLeft;
}
