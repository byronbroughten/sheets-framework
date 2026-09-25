import {
  type CellValueName,
  type UniformRowName,
} from "../00_Source/CellValues/cellValues";
import { dimensionIds } from "../01_SpreadsheetSchema/dimensionIds";
import { Val } from "../utils/Val";
import { SheetCommonRaw } from "./ClassBases/SheetCommonRaw";
import { ColumnMetaRaw } from "./ColumnMetaRaw";
import { SheetRaw } from "./SheetRaw";
import { SpreadsheetRaw } from "./SpreadsheetRaw";
import { UniformRowRaw } from "./UniformRowRaw";

export class SheetMetaRaw extends SheetCommonRaw {
  get ss(): SpreadsheetRaw {
    return new SpreadsheetRaw(this.spreadsheetRawProps);
  }
  get primary(): SheetRaw {
    return new SheetRaw(this.sheetRawProps);
  }
  get hasFetchedColumnIds(): boolean {
    return this.sheetState.working.hasFetchedColumnIds;
  }
  get tableHeaderRow(): UniformRowRaw<"tableHeader"> {
    return this.uniformRow("tableHeader");
  }
  get actionRow(): UniformRowRaw<"action"> {
    return this.uniformRow("action");
  }
  get colIdRow(): UniformRowRaw<"columnId"> {
    return this.uniformRow("columnId");
  }
  get activeColumnIds(): string[] {
    return this._tableColumnIds().filter((columnId) => columnId !== "");
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
  columnIdAt(colIndex: number): string {
    if (this.isTableColIndex(colIndex)) {
      return this._columnIdInTable(colIndex);
    }
    const value = this.colIdRow.valueOrEmpty(colIndex);
    return typeof value === "string" ? value : "";
  }
  uniformRow<UN extends UniformRowName>(uniformRowName: UN): UniformRowRaw<UN> {
    return new UniformRowRaw({
      ...this.sheetRawProps,
      uniformRowName,
    });
  }
  uniformRowByIndex(rowIndex: number): UniformRowRaw {
    return this.uniformRow(this.schema.uniformRowNameByIndex(rowIndex));
  }
  column<VN extends CellValueName = CellValueName>(
    colIndex: number,
  ): ColumnMetaRaw<VN> {
    return new ColumnMetaRaw<VN>({
      colIndex,
      ...this.sheetRawProps,
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
        `Value ${columnId} not found in row ${this.schema.colIdRowIndex}. Cannot find column index.`,
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
  insertColumnAtEnd(props: { idPrefix: string; header: string }): number {
    const columnIndex = this.nextEndColumnInsertIndex;
    this.addSheetChangeToSave({
      action: "insertColumn",
      startColumnIndex: columnIndex,
    });
    this.column(columnIndex).initUniformCells(props);
    return columnIndex;
  }
  // Only table columns: a fact is always reached through a column ID.
  ensureTableColumnsActiveFacts(): void {
    this.fullTableColIndexes.forEach((colIndex) => {
      this.column(colIndex).ensureActiveFacts();
    });
  }
  gatherFetchColumnIdsInit(startTableColIndex: number): this {
    this.gatherFetchRange({
      startRowIndex: this.schema.colIdRowIndex,
      endRowIndex: this.schema.colIdRowIndex + 1,
      startColumnIndex: startTableColIndex,
    });
    this.sheetState.fetchQueue.toFinalize.rows.add(this.schema.colIdRowIndex);
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
