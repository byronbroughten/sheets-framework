import { dimensionIds } from "../01_SpreadsheetSchema/dimensionIds";
import { headRows } from "../01_SpreadsheetSchema/headRows";
import { SpreadsheetSchema } from "../01_SpreadsheetSchema/SpreadsheetSchema";
import { getTableTraitByName } from "../01_SpreadsheetSchema/tableConfigsTypes";
import { TableOrigin } from "../01_SpreadsheetSchema/TableOrigin";
import {
  buildGridRows,
  type FakeRichCellValue,
  type FakeSheetProperties,
  fakeTableId,
  type stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { SpreadsheetRaw } from "./SpreadsheetRaw";

export const lightGreen = { red: 0.851, green: 0.918, blue: 0.827 };

export const itemGid = getTableTraitByName("item", "sheetGid");
export const logGid = getTableTraitByName("log", "sheetGid");
export const itemTableId = fakeTableId(itemGid, 0);
export const logTableId = fakeTableId(logGid, 0);
// Sheet rows and columns, for fixtures and grid reads; Raw itself counts from the Table.
export const expectedOrigin = TableOrigin.expected();
export const tableHeaderRowIndex: number = expectedOrigin.headerRowIndex;
export const colIdRowIndex: number = expectedOrigin.sheetRowIndex(
  headRows.index("columnId"),
);
export const startTableColIndex: number = expectedOrigin.startColIndex;
export const topDataRowIndex = tableHeaderRowIndex + 1;
export const scratchGid = 999999;
export const scratchTableId = fakeTableId(scratchGid, 0);
export const tableId111 = fakeTableId(111, 0);
export const tableId222 = fakeTableId(222, 0);
export const tableEndRowIndex = tableHeaderRowIndex + 3;
export const gridRanges = {
  columnOneData: {
    startRowIndex: topDataRowIndex,
    startColumnIndex: 1,
    endColumnIndex: 2,
  },
  headerAndTopDataRow: {
    startRowIndex: tableHeaderRowIndex,
    endRowIndex: topDataRowIndex + 1,
  },
} as const;

export function threeByThreeAround(
  rowIndex: number,
  colIndex: number,
): {
  startRowIndex: number;
  endRowIndex: number;
  startColumnIndex: number;
  endColumnIndex: number;
} {
  return {
    startRowIndex: rowIndex - 1,
    endRowIndex: rowIndex + 2,
    startColumnIndex: colIndex - 1,
    endColumnIndex: colIndex + 2,
  };
}

export function firstTableEndRowIndex(
  grid: ReturnType<typeof stubSheetsService>["grid"],
  sheetGid: number,
): number | undefined {
  return grid.sheet(sheetGid).tables[0]?.range?.endRowIndex;
}

export function placedTableSheet(sheet: {
  sheetId: number;
  title: string;
}): FakeSheetProperties {
  return {
    ...sheet,
    rows: buildGridRows({
      [colIdRowIndex]: [ownColumnId(sheet.sheetId)],
      [tableHeaderRowIndex]: ["ID"],
    }),
    table: { endRowIndex: tableEndRowIndex },
  };
}

// A sheet the config doesn't know has no prefix of its own, so any well-formed ID does.
export function ownColumnId(sheetGid: number): string {
  const schema = new SpreadsheetSchema();
  if (!schema.isInSheetGids(sheetGid)) return dimensionIds.col("x", "id");
  return dimensionIds.col(schema.sheetByGid(sheetGid).idPrefix, "id");
}

export function misplacedTableSheet({
  startRowIndex = tableHeaderRowIndex,
  startColumnIndex = startTableColIndex,
  ...sheet
}: {
  sheetId: number;
  title: string;
  startRowIndex?: number;
  startColumnIndex?: number;
}): FakeSheetProperties {
  return {
    ...placedTableSheet(sheet),
    table: {
      endRowIndex: tableEndRowIndex,
      startRowIndex,
      startColumnIndex,
    },
  };
}

export function extraTablesSheet(sheet: {
  sheetId: number;
  title: string;
}): FakeSheetProperties {
  return {
    ...sheet,
    rows: buildGridRows({ [tableHeaderRowIndex]: ["ID"] }),
    tables: [
      { endRowIndex: tableEndRowIndex },
      {
        startRowIndex: tableHeaderRowIndex + 10,
        endRowIndex: tableHeaderRowIndex + 12,
      },
    ],
  };
}

export function recordedGridRanges(calls: object[]): unknown[] {
  const resource = calls[0] as {
    dataFilters: { gridRange: unknown }[];
  };
  return resource.dataFilters.map((filter) => filter.gridRange);
}

export function formulaCell(formula: string): FakeRichCellValue {
  return { value: formula, isFormula: true };
}

export function fetchedRaw(): SpreadsheetRaw {
  const raw = SpreadsheetRaw.init();
  raw.fetchAllSheetProperties();
  return raw;
}

export function thrownMessage(fn: () => void): string {
  try {
    fn();
  } catch (error) {
    return (error as Error).message;
  }
  throw new Error("Expected the call to throw, but it did not.");
}
