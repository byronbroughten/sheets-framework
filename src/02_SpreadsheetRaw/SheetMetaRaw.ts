import { dimensionIds } from "../01_SpreadsheetSchema/dimensionIds";
import { TableCommonRaw } from "./ClassBases/TableCommonRaw";
import type { TableEndColumnHeadCells } from "./ClassTypes/StateRaw";
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
