import { PartialTableRefusal } from "../../00_Source/RawSource/PartialTableRefusal";
import type {
  InsertRangeOperation,
  TableSnapshot,
} from "../../00_Source/RawSource/RawSource";
import { SheetIndex } from "../../00_Source/RawSource/SheetIndex";
import { SpreadsheetBaseRaw } from "../ClassBases/SpreadsheetBaseRaw";
import { rowShiftFrom } from "../ClassBases/TableCommonRaw";
import type { AppendTableRows } from "../ClassTypes/StateRaw";
import { SpreadsheetRaw } from "../SpreadsheetRaw";

export class SpreadsheetRefusalRewriterRaw extends SpreadsheetBaseRaw {
  get ss(): SpreadsheetRaw {
    return new SpreadsheetRaw(this.spreadsheetRawProps);
  }
  // Google names neither Table, so the Table refused is the one the operation was gathered for.
  rewordedRefusal(refusal: PartialTableRefusal): PartialTableRefusal {
    const { operation } = refusal;
    if (operation.kind !== "insertRange") return refusal;
    const split = this._fetchTableSplitBy(operation);
    if (split === undefined) return refusal;
    const growth = this.writeOperations.appendTableRows.find(({ operations }) =>
      operations.includes(operation),
    );
    if (growth !== undefined) {
      const grown = this.ss.table(growth.tableId);
      return new PartialTableRefusal(
        `Growing Table "${grown.name}" on sheet "${grown.sheetTitle}" would insert cells over part of Table "${split.name}" below it. Make the lower Table, "${split.name}", no wider than "${grown.name}", or move it.`,
        operation,
      );
    }
    const splitting = this._columnWriteSplitting(operation, split.name);
    if (splitting === undefined) return refusal;
    const { startRowIndex, endRowIndex } = operation.range;
    return new PartialTableRefusal(
      `${splitting}. Move "${split.name}" so that it sits entirely within rows ${startRowIndex + 1}–${endRowIndex}, or entirely outside them.`,
      operation,
    );
  }
  // Refused whole, so the live layout is from before this batch's growth, which a column insert is measured after.
  private _fetchTableSplitBy(
    insert: InsertRangeOperation,
  ): TableSnapshot | undefined {
    const { sheetId } = insert.range;
    const sheet = this.spreadsheetStateRaw.rawSource
      .fetchSheetProperties()
      .sheets.find(({ sheetGid }) => sheetGid === sheetId);
    const growths = this.writeOperations.appendTableRows;
    return sheet?.tables
      ?.map((table) => shiftedDownBy(growths, { ...table, sheetId }))
      .find((table) => isPartlyShiftedBy(insert, table));
  }
  private _columnWriteSplitting(
    insert: InsertRangeOperation,
    splitName: string,
  ): string | undefined {
    const { insertTableEndColumns, changeTableColumns } = this.writeOperations;
    const endInsert = insertTableEndColumns.find(({ operations }) =>
      operations.includes(insert),
    );
    if (endInsert !== undefined) {
      return this.ss.table(endInsert.tableId).columnInsertSplitting(splitName);
    }
    const change = changeTableColumns.find(({ operations }) =>
      operations.includes(insert),
    );
    if (change === undefined) return undefined;
    return this.ss
      .table(change.tableId)
      .columnChangeSplitting(change, insert, splitName);
  }
}

function shiftedDownBy(
  growths: AppendTableRows[],
  table: TableSnapshot & { sheetId: number },
): TableSnapshot {
  const rowCount = rowShiftFrom(growths, table);
  return {
    ...table,
    startRowIndex: SheetIndex.row(table.startRowIndex + rowCount),
    endRowIndex: SheetIndex.row(table.endRowIndex + rowCount),
  };
}

// Google's rule: an insert reaching a Table must span all of it across the shift.
function isPartlyShiftedBy(
  { range, shiftDimension }: InsertRangeOperation,
  table: TableSnapshot,
): boolean {
  if (shiftDimension === "ROWS") {
    return (
      table.endRowIndex > range.startRowIndex &&
      coversPartOf(
        [range.startColumnIndex, range.endColumnIndex],
        [table.startColumnIndex, table.endColumnIndex],
      )
    );
  }
  return (
    table.endColumnIndex > range.startColumnIndex &&
    coversPartOf(
      [range.startRowIndex, range.endRowIndex],
      [table.startRowIndex, table.endRowIndex],
    )
  );
}

function coversPartOf(
  [bandStart, bandEnd]: [number, number],
  [start, end]: [number, number],
): boolean {
  const overlaps = bandStart < end && start < bandEnd;
  const coversAll = bandStart <= start && end <= bandEnd;
  return overlaps && !coversAll;
}
