import { dimensionIds } from "../../01_SpreadsheetSchema/dimensionIds";
import { Val } from "../../utils/Val";
import { TableCommonRaw } from "../ClassBases/TableCommonRaw";
import { SpreadsheetRaw } from "../SpreadsheetRaw";
import { TableRaw } from "../TableRaw";

export class TableColumnResolverRaw extends TableCommonRaw {
  get ss(): SpreadsheetRaw {
    return new SpreadsheetRaw(this.spreadsheetRawProps);
  }
  get table(): TableRaw {
    return new TableRaw(this.tableRawProps);
  }
  get hasFetchedColumnIds(): boolean {
    return this.tableState.working.hasFetchedColumnIds;
  }
  colIndexOfHeader(header: string): number {
    return this.table.headRow("header").colIndexOfValue(header);
  }
  columnIdByHeader(header: string): string {
    return this.columnIdAt(this.colIndexOfHeader(header));
  }
  columnIdAt(colIndex: number): string {
    if (this.isTableColIndex(colIndex)) {
      return this._columnIdInTable(colIndex);
    }
    const value = this.table.headRow("columnId").valueOrEmpty(colIndex);
    return typeof value === "string" ? value : "";
  }
  get colIndexesWithoutColumnId(): number[] {
    return this.fullTableColIndexes.filter(
      (colIndex) => this._columnIdInTable(colIndex) === "",
    );
  }
  hasColumnId(columnId: string): boolean {
    return this._tableColumnIds().includes(columnId);
  }
  colIndexOf(columnId: string): number {
    const tableColIndexes = this.fullTableColIndexes;
    const colIndex = this._tableColumnIds().findIndex((id) => id === columnId);
    if (colIndex === -1) {
      throw new Error(
        `Value ${columnId} not found in ${this.rowLabel(this.schema.colIdRowIndex)}. Cannot find column index.`,
      );
    }
    return Val.assert(tableColIndexes[colIndex], "Table column index");
  }
  // Fetched cells only: the placement strip carries just the first column's.
  holdsOnlyColumnIdsOf(idPrefix: string): boolean {
    const columnIds = this.fullTableColIndexes
      .map((colIndex) => this.table.headRow("columnId").cell(colIndex))
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
  gatherFetchColumnIds(): this {
    const colIdRowIndex = this.schema.colIdRowIndex;
    this.gatherFetchRange(this.fullRowFetchRange(colIdRowIndex));
    this.tableState.fetchQueue.toFinalize.rows.add(colIdRowIndex);
    return this;
  }
  private _columnIdInTable(colIndex: number): string {
    const value = this.table.headRow("columnId").valueOrEmpty(colIndex);
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
