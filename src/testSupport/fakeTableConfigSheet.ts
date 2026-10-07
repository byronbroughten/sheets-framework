import { getColumnTraitByName } from "../01_SpreadsheetSchema/columnConfigsTypes";
import { floorTabSeedByGid } from "../01_SpreadsheetSchema/configSheetFloorSeed";
import {
  getTableTraitByGid,
  getTableTraitByName,
} from "../01_SpreadsheetSchema/tableConfigsTypes";
import {
  buildGridRows,
  type FakeCell,
  type FakeSheetsService,
  fakeTableId,
  stubSheetsService,
} from "./fakeSheetsService";

/**
 * A fake "Table Config" sheet, for tests about behaviour that reads or writes
 * a whole row rather than one named cell — clearing, the blank test, the wipe,
 * append reuse. Table Config earns the job by being the smallest sheet every config set shares.
 *
 * The column ids are the installed configs' ones, and the default columnId row
 * lists every column the config declares so clear / blank / wipe resolve each.
 */
export const tableConfigGid = getTableTraitByName("tableConfig", "sheetGid");

export const tableConfigColumnIdRow = (
  ["tableId", "tableName", "sheetTitle", "letApiAccess"] as const
).map((columnName) =>
  getColumnTraitByName("tableConfig", columnName, "columnId"),
);

export const sheetTitleColIndex = tableConfigColumnIdRow.indexOf(
  getColumnTraitByName("tableConfig", "sheetTitle", "columnId"),
);

/** The tableId a tab's one Table carries: a floor tab's generated one, else the fake's default. */
export function tableIdOnTab(sheetGid: number): string {
  if (floorTabSeedByGid(sheetGid) === undefined)
    {return fakeTableId(sheetGid, 0);}
  return getTableTraitByGid(sheetGid, "tableId");
}

/** Every non-formula column filled in. */
export const filledTableConfigRow: FakeCell[] = [
  "fake-table-999001",
  "item",
  "Item",
  true,
];

/** The blank row: nothing in any non-formula column. */
export const blankTableConfigRow: FakeCell[] = [null, null, null, null];

/**
 * Stubs the Sheets service with just this sheet, its data rows keyed by
 * literal sheet row index (4 is the top data row). Pass `rowsWithNoGridData`
 * to make a row come back with no cell data, which is what a row nothing was
 * ever written to looks like.
 */
export function stubTableConfigSheet(
  dataRows: Record<number, FakeCell[]>,
  rowsWithNoGridData: number[] = [],
): FakeSheetsService {
  const rowIndexes = Object.keys(dataRows).map(Number);
  return stubSheetsService({
    sheets: [
      {
        sheetId: tableConfigGid,
        title: "Table Config",
        rows: buildGridRows({ 0: tableConfigColumnIdRow, ...dataRows }),
        rowsWithNoGridData,
        table: { endRowIndex: Math.max(...rowIndexes) + 1 },
      },
    ],
  });
}
