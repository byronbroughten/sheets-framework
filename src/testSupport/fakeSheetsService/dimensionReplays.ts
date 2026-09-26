import type { FakeCell, FakeCellValue } from "../fakeSheetsService";
import { fakeCells } from "./fakeCells";
import { type DimensionChange, fakeGrid } from "./fakeGrid";
import { type FakeSpreadsheet, fakeSpreadsheet } from "./fakeSpreadsheet";

type Response = GoogleAppsScript.Sheets.Schema.Response;
type DimensionRange = GoogleAppsScript.Sheets.Schema.DimensionRange;

// Measured live: ascending is numbers, text, then booleans; descending reverses it.
const sortTypeRanks = { number: 0, string: 1, boolean: 2 } as const;

export const dimensionReplays = {
  // Measured live: a Table append inserts rows at the Table's end, shifting rows below.
  appendCells(
    spreadsheet: FakeSpreadsheet,
    request: GoogleAppsScript.Sheets.Schema.AppendCellsRequest,
  ): Response {
    const fields = fakeCells.fields(request.fields, "appendCells");
    const rows = request.rows ?? [];
    const sheet = fakeSpreadsheet.sheet(spreadsheet, request.sheetId);
    let startRowIndex = fakeGrid.endOfData(sheet);
    let startColumnIndex = 0;
    if (request.tableId !== undefined) {
      const { table } = fakeSpreadsheet.table(spreadsheet, request.tableId);
      startRowIndex = table.endRowIndex;
      startColumnIndex = table.startColumnIndex;
      fakeGrid.insert(
        sheet,
        { dimension: "ROWS", startIndex: startRowIndex, count: rows.length },
        true,
      );
    }
    rows.forEach((row, rowOffset) => {
      const rowIndex = startRowIndex + rowOffset;
      sheet.rowCount = Math.max(sheet.rowCount, rowIndex + 1);
      (row.values ?? []).forEach((data, colOffset) => {
        const colIndex = startColumnIndex + colOffset;
        fakeGrid.updateCell(sheet, rowIndex, colIndex, (cell) =>
          fakeCells.withCellData(cell, data, fields),
        );
      });
    });
    return {};
  },
  insertDimension(
    spreadsheet: FakeSpreadsheet,
    request: GoogleAppsScript.Sheets.Schema.InsertDimensionRequest,
  ): Response {
    fakeGrid.insert(
      fakeSpreadsheet.sheet(spreadsheet, request.range?.sheetId),
      dimensionChange(request.range),
      request.inheritFromBefore === true,
    );
    return {};
  },
  deleteDimension(
    spreadsheet: FakeSpreadsheet,
    request: GoogleAppsScript.Sheets.Schema.DeleteDimensionRequest,
  ): Response {
    fakeGrid.remove(
      fakeSpreadsheet.sheet(spreadsheet, request.range?.sheetId),
      dimensionChange(request.range),
    );
    return {};
  },
  updateDimensionProperties(
    spreadsheet: FakeSpreadsheet,
    request: GoogleAppsScript.Sheets.Schema.UpdateDimensionPropertiesRequest,
  ): Response {
    if (request.fields !== "hiddenByUser") {
      throw new Error(
        `The fake Sheets service does not replay updateDimensionProperties fields "${request.fields}".`,
      );
    }
    const { dimension, startIndex, count } = dimensionChange(request.range);
    const sheet = fakeSpreadsheet.sheet(spreadsheet, request.range?.sheetId);
    const changed = Array.from({ length: count }, (_, i) => startIndex + i);
    const hidden = new Set(
      dimension === "ROWS" ? sheet.hiddenRowIndexes : sheet.hiddenColumnIndexes,
    );
    changed.forEach((index) => {
      if (request.properties?.hiddenByUser === true) {
        hidden.add(index);
      } else {
        hidden.delete(index);
      }
    });
    const sorted = [...hidden].sort((a, b) => a - b);
    if (dimension === "ROWS") {
      sheet.hiddenRowIndexes = sorted;
    } else {
      sheet.hiddenColumnIndexes = sorted;
    }
    return {};
  },
  // Measured live: blanks sort last whichever the order.
  sortRange(
    spreadsheet: FakeSpreadsheet,
    request: GoogleAppsScript.Sheets.Schema.SortRangeRequest,
  ): Response {
    const sheet = fakeSpreadsheet.sheet(spreadsheet, request.range?.sheetId);
    const range = fakeGrid.boundedRange(sheet, request.range ?? {});
    const specs = request.sortSpecs ?? [];
    const rowIndexes = Array.from(
      { length: range.endRowIndex - range.startRowIndex },
      (_, i) => range.startRowIndex + i,
    );
    const colIndexes = Array.from(
      { length: range.endColumnIndex - range.startColumnIndex },
      (_, i) => range.startColumnIndex + i,
    );
    const rows = rowIndexes.map((rowIndex) =>
      colIndexes.map((colIndex) => fakeGrid.cell(sheet, rowIndex, colIndex)),
    );
    const sorted = [...rows].sort((a, b) => {
      for (const spec of specs) {
        const keyIndex = (spec.dimensionIndex ?? 0) - range.startColumnIndex;
        const order = compareSortKeys(
          a[keyIndex] ?? null,
          b[keyIndex] ?? null,
          spec.sortOrder === "DESCENDING",
        );
        if (order !== 0) return order;
      }
      return 0;
    });
    sorted.forEach((row, rowOffset) => {
      row.forEach((cell, colOffset) => {
        fakeGrid.setCell(
          sheet,
          range.startRowIndex + rowOffset,
          range.startColumnIndex + colOffset,
          cell,
        );
      });
    });
    return {};
  },
};

function dimensionChange(range: DimensionRange | undefined): DimensionChange {
  const dimension = range?.dimension;
  if (dimension !== "ROWS" && dimension !== "COLUMNS") {
    throw new Error(
      `A dimension range needs ROWS or COLUMNS, not ${dimension}.`,
    );
  }
  const startIndex = range?.startIndex ?? 0;
  const endIndex = range?.endIndex;
  if (endIndex === undefined || endIndex <= startIndex) {
    throw new Error(
      `A dimension range needs an end past its start (${startIndex}, ${endIndex}).`,
    );
  }
  return {
    dimension,
    startIndex,
    count: endIndex - startIndex,
  };
}

function compareSortKeys(
  a: FakeCell,
  b: FakeCell,
  isDescending: boolean,
): number {
  const aIsBlank = fakeCells.isBlank(a);
  const bIsBlank = fakeCells.isBlank(b);
  if (aIsBlank || bIsBlank) return Number(aIsBlank) - Number(bIsBlank);
  const order = compareValues(fakeCells.value(a), fakeCells.value(b));
  return isDescending ? -order : order;
}

function compareValues(a: FakeCellValue, b: FakeCellValue): number {
  if (a === null || b === null) return 0;
  const typeOrder =
    sortTypeRanks[typeof a as keyof typeof sortTypeRanks] -
    sortTypeRanks[typeof b as keyof typeof sortTypeRanks];
  if (typeOrder !== 0) return typeOrder;
  if (typeof a === "string" && typeof b === "string") return a.localeCompare(b);
  return Number(a) - Number(b);
}
