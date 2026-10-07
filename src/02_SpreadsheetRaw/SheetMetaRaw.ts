import { dimensionIds } from "../01_SpreadsheetSchema/dimensionIds";
import { TableCommonRaw } from "./ClassBases/TableCommonRaw";
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
}
