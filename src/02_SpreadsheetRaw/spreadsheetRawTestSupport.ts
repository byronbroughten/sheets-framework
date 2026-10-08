import { dimensionIds } from "../01_SpreadsheetSchema/dimensionIds";
import { headRows } from "../01_SpreadsheetSchema/headRows";
import { SpreadsheetSchema } from "../01_SpreadsheetSchema/SpreadsheetSchema";
import { getTableTraitByName } from "../01_SpreadsheetSchema/tableConfigsTypes";
import { TableOrigin } from "../01_SpreadsheetSchema/TableOrigin";
import {
  buildGridRows,
  type FakeCell,
  type FakeRichCellValue,
  type FakeSheetProperties,
  type FakeTable,
  fakeTableId,
  type stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { Val } from "../utils/Val";
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

export const layoutGid = getTableTraitByName("layoutLeft", "sheetGid");
export const layoutTableNames = [
  "layoutLeft",
  "layoutRight",
  "layoutBelow",
] as const;
export type LayoutTableName = (typeof layoutTableNames)[number];
export function layoutTableId(tableName: LayoutTableName): string {
  return getTableTraitByName(tableName, "tableId");
}
export const layoutBodyRows: Record<LayoutTableName, [string, number][]> = {
  layoutLeft: [
    ["Left one", 1],
    ["Left two", 2],
    ["Left three", 3],
  ],
  layoutRight: [
    ["Right one", 10],
    ["Right two", 20],
  ],
  layoutBelow: [
    ["Below one", 100],
    ["Below two", 200],
  ],
};

// The dev Layout sheet: Right beside Left, Below under Left, each at its recorded origin unless overridden.
export function layoutSheet(
  overrides: Partial<Record<LayoutTableName, Partial<FakeTable>>> = {},
): FakeSheetProperties {
  const placedTables = layoutTableNames.map((tableName) => ({
    tableName,
    table: layoutTable(tableName, overrides[tableName] ?? {}),
  }));
  return {
    sheetId: layoutGid,
    title: "Layout",
    rows: gridRowsOf(placedTables.flatMap(headerAndBodyCells)),
    tables: placedTables.map(({ table }) => table),
  };
}

function layoutTable(
  tableName: LayoutTableName,
  override: Partial<FakeTable>,
): FakeTable {
  const startRowIndex =
    override.startRowIndex ?? getTableTraitByName(tableName, "headerRowIndex");
  const startColumnIndex =
    override.startColumnIndex ??
    getTableTraitByName(tableName, "startColIndex");
  return {
    tableId: layoutTableId(tableName),
    name: tableName,
    headRows: {
      3: ["entry", "amount"].map((columnKey) =>
        dimensionIds.col(getTableTraitByName(tableName, "idPrefix"), columnKey),
      ),
    },
    endRowIndex: startRowIndex + 1 + layoutBodyRows[tableName].length,
    endColumnIndex: startColumnIndex + 2,
    ...override,
    startRowIndex,
    startColumnIndex,
  };
}

interface PlacedCell {
  rowIndex: number;
  colIndex: number;
  cell: FakeCell;
}

function headerAndBodyCells({
  tableName,
  table,
}: {
  tableName: LayoutTableName;
  table: FakeTable;
}): PlacedCell[] {
  const headerRowIndex = Val.assert(table.startRowIndex, "layout header row");
  const startColIndex = Val.assert(
    table.startColumnIndex,
    "layout start column",
  );
  const headerAndBody: readonly (readonly FakeCell[])[] = [
    ["Entry", "Amount"],
    ...layoutBodyRows[tableName],
  ];
  return headerAndBody.flatMap((cells, rowOffset) =>
    cells.map((cell, colOffset) => ({
      rowIndex: headerRowIndex + rowOffset,
      colIndex: startColIndex + colOffset,
      cell,
    })),
  );
}

function gridRowsOf(cells: PlacedCell[]): FakeCell[][] {
  return cells.reduce<FakeCell[][]>((rows, { rowIndex, colIndex, cell }) => {
    while (rows.length <= rowIndex) rows.push([]);
    const row = Val.assert(rows[rowIndex], `layout row ${rowIndex}`);
    while (row.length < colIndex) row.push(null);
    row[colIndex] = cell;
    return rows;
  }, []);
}
