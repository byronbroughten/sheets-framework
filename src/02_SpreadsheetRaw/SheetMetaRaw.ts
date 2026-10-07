import type { CellValueName } from "../00_Source/CellValues/cellValues";
import { dimensionIds } from "../01_SpreadsheetSchema/dimensionIds";
import { Val } from "../utils/Val";
import { TableCommonRaw } from "./ClassBases/TableCommonRaw";
import type { TableEndColumnHeadCells } from "./ClassTypes/StateRaw";
import { ColumnMetaRaw } from "./ColumnMetaRaw";
import type { HeadRowRaw } from "./HeadRowRaw";
import { SpreadsheetRaw } from "./SpreadsheetRaw";
import { TableRaw } from "./TableRaw";

export class SheetMetaRaw extends TableCommonRaw {
  get ss(): SpreadsheetRaw {
    return new SpreadsheetRaw(this.spreadsheetRawProps);
  }
  get primary(): TableRaw {
    return new TableRaw(this.tableRawProps);
  }
  get hasFetchedColumnIds(): boolean {
    return this.tableState.working.hasFetchedColumnIds;
  }
  get tableHeaderRow(): HeadRowRaw<"header"> {
    return this.primary.headRow("header");
  }
  get actionRow(): HeadRowRaw<"action"> {
    return this.primary.headRow("action");
  }
  get colIdRow(): HeadRowRaw<"columnId"> {
    return this.primary.headRow("columnId");
  }
  get activeColumnIds(): string[] {
    return this._tableColumnIds().filter((columnId) => columnId !== "");
  }
  // Fetched cells only: the placement strip carries just the first column's.
  holdsOnlyColumnIdsOf(idPrefix: string): boolean {
    const columnIds = this.fullTableColIndexes
      .map((colIndex) => this.colIdRow.cell(colIndex))
      .filter((cell) => cell.inWorking)
      .map((cell) => cell.valueOrEmpty())
      .filter((value) => value !== "");
    return (
      columnIds.length > 0 &&
      columnIds.every(
        (columnId) =>
          typeof columnId === "string" &&
          dimensionIds.colIdPrefixOrUndefined(columnId) === idPrefix,
      )
    );
  }
  activeIdPrefix(): string | undefined {
    const columnIdsByPrefix = this._columnIdsByIdPrefix();
    if (columnIdsByPrefix.size === 0) return undefined;
    if (columnIdsByPrefix.size > 1) {
      throw new Error(
        mixedIdPrefixMessage(this.primary.title, columnIdsByPrefix),
      );
    }
    return columnIdsByPrefix.keys().next().value;
  }
  private _columnIdsByIdPrefix(): Map<string, string[]> {
    const columnIdsByPrefix = new Map<string, string[]>();
    this.activeColumnIds.forEach((columnId) => {
      const idPrefix = dimensionIds.colIdPrefixOrUndefined(columnId);
      if (idPrefix === undefined) return;
      const columnIds = columnIdsByPrefix.get(idPrefix) ?? [];
      columnIds.push(columnId);
      columnIdsByPrefix.set(idPrefix, columnIds);
    });
    return columnIdsByPrefix;
  }
  columnIdByHeader(header: string): string {
    return this.columnIdAt(this.tableHeaderRow.colIndexOfValue(header));
  }
  columnIdAt(colIndex: number): string {
    if (this.isTableColIndex(colIndex)) {
      return this._columnIdInTable(colIndex);
    }
    const value = this.colIdRow.valueOrEmpty(colIndex);
    return typeof value === "string" ? value : "";
  }
  column<VN extends CellValueName = CellValueName>(
    colIndex: number,
  ): ColumnMetaRaw<VN> {
    return new ColumnMetaRaw<VN>({
      colIndex,
      ...this.tableRawProps,
    });
  }
  columnByActiveId<VN extends CellValueName = CellValueName>(
    columnId: string,
  ): ColumnMetaRaw<VN> {
    return this.column<VN>(this.colIndexOfActiveColumnId(columnId));
  }
  isActiveColumnId(columnId: string): boolean {
    return this._tableColumnIds().includes(columnId);
  }
  colIndexOfActiveColumnId(columnId: string): number {
    const tableColIndexes = this.fullTableColIndexes;
    const colIndex = this._tableColumnIds().findIndex((id) => id === columnId);
    if (colIndex === -1) {
      throw new Error(
        `Value ${columnId} not found in ${this.rowLabel(this.schema.colIdRowIndex)}. Cannot find column index.`,
      );
    }
    return Val.assert(tableColIndexes[colIndex], "Table column index");
  }
  addMissingColumnIds(idPrefix: string): number {
    let addedCount = 0;
    this.fullTableColIndexes.forEach((colIndex) => {
      const colIdValue = this._columnIdInTable(colIndex);
      if (!colIdValue) {
        this.colIdRow.updateValue(colIndex, dimensionIds.col(idPrefix));
        addedCount++;
      }
    });
    return addedCount;
  }
  // Past the inserts already queued, since each lands at the Table end as it stands then.
  insertColumnAtEnd(headCells: TableEndColumnHeadCells): number {
    const colIndex = this.columnCount + this.writes.insertTableEndColumnCount;
    this.queueTableWrite({ action: "insertTableEndColumn" });
    this.column(colIndex).initHeadCells(headCells);
    return colIndex;
  }
  // Only table columns: a fact is always reached through a column ID.
  ensureTableColumnsActiveFacts(): void {
    this.fullTableColIndexes.forEach((colIndex) => {
      this.column(colIndex).ensureActiveFacts();
    });
  }
  gatherFetchColumnIdsInit(): this {
    const colIdRowIndex = this.schema.colIdRowIndex;
    this.gatherFetchRange(this.fullRowFetchRange(colIdRowIndex));
    this.tableState.fetchQueue.toFinalize.rows.add(colIdRowIndex);
    return this;
  }
  private _columnIdInTable(colIndex: number): string {
    const value = this.colIdRow.valueOrEmpty(colIndex);
    if (value !== "" && typeof value !== "string") {
      throw new Error(
        `Column ID in ${this.sheetLabel} column index ${colIndex} must be text or blank, got ${JSON.stringify(value)}.`,
      );
    }
    return value;
  }
  private _tableColumnIds(): string[] {
    return this.fullTableColIndexes.map((colIndex) =>
      this._columnIdInTable(colIndex),
    );
  }
}

function mixedIdPrefixMessage(
  sheetTitle: string,
  columnIdsByPrefix: Map<string, string[]>,
): string {
  const prefixParts = [...columnIdsByPrefix.entries()].map(
    ([idPrefix, columnIds]) =>
      `"${idPrefix}" (${columnIds.length} column${
        columnIds.length === 1 ? "" : "s"
      }: ${columnIds.join(", ")})`,
  );
  return (
    `Sheet "${sheetTitle}" has column IDs with more than one ID prefix: ` +
    `${prefixParts.join("; ")}. Clear the stray column ID cells so the next ` +
    `sync can mint new ones.`
  );
}
