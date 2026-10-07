import type { CellValueName } from "../00_Source/CellValues/cellValues";
import { dimensionIds } from "../01_SpreadsheetSchema/dimensionIds";
import { TableCommonRaw } from "./ClassBases/TableCommonRaw";
import type { TableEndColumnHeadCells } from "./ClassTypes/StateRaw";
import type { ColumnMetaRaw } from "./ColumnMetaRaw";
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
  get actionRow(): HeadRowRaw<"action"> {
    return this.primary.headRow("action");
  }
  get activeColumnIds(): string[] {
    return this.fullTableColIndexes
      .map((colIndex) => this.primary.columnResolver.columnIdAt(colIndex))
      .filter((columnId) => columnId !== "");
  }
  // Fetched cells only: the placement strip carries just the first column's.
  holdsOnlyColumnIdsOf(idPrefix: string): boolean {
    const columnIds = this.fullTableColIndexes
      .map((colIndex) => this.primary.headRow("columnId").cell(colIndex))
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
  columnByActiveId<VN extends CellValueName = CellValueName>(
    columnId: string,
  ): ColumnMetaRaw<VN> {
    return this.primary.column<VN>(
      this.primary.columnResolver.colIndexOf(columnId),
    ).meta;
  }
  addMissingColumnIds(idPrefix: string): number {
    let addedCount = 0;
    this.fullTableColIndexes.forEach((colIndex) => {
      const colIdValue = this.primary.columnResolver.columnIdAt(colIndex);
      if (!colIdValue) {
        this.primary
          .headRow("columnId")
          .updateValue(colIndex, dimensionIds.col(idPrefix));
        addedCount++;
      }
    });
    return addedCount;
  }
  // Past the inserts already queued, since each lands at the Table end as it stands then.
  insertColumnAtEnd(headCells: TableEndColumnHeadCells): number {
    const colIndex = this.columnCount + this.writes.insertTableEndColumnCount;
    this.queueTableWrite({ action: "insertTableEndColumn" });
    this.primary.column(colIndex).meta.initHeadCells(headCells);
    return colIndex;
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
