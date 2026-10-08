import { SheetIndex } from "../../00_Source/RawSource/SheetIndex";
import { TableOrigin } from "../../01_SpreadsheetSchema/TableOrigin";
import type { FakeCell, FakeTable } from "../fakeSheetsService";
import { fakeCells } from "./fakeCells";
import { type FakeSheetState, type FakeTableState } from "./fakeSpreadsheet";

export type FakeTablePlacement = Pick<
  FakeTable,
  "startRowIndex" | "startColumnIndex"
>;

type GridRange = GoogleAppsScript.Sheets.Schema.GridRange;
type Table = GoogleAppsScript.Sheets.Schema.Table;
type TableColumnProperties =
  GoogleAppsScript.Sheets.Schema.TableColumnProperties;

export const fakeTables = {
  // Google's shape for a sheet's Tables, as `get` returns them and the grid view shows them.
  googleTables(sheet: FakeSheetState): Table[] | undefined {
    return toGoogleTables(sheet, sheet.tables);
  },
  // A filtered fetch returns only the Tables its ranges overlap.
  googleTablesOverlapping(
    sheet: FakeSheetState,
    filterRanges: readonly GridRange[],
  ): Table[] | undefined {
    const tables = sheet.tables.filter((table) =>
      filterRanges.some(
        (range) => range.sheetId === sheet.sheetId && overlaps(table, range),
      ),
    );
    return toGoogleTables(sheet, tables);
  },
  origin({ startRowIndex, startColumnIndex }: FakeTablePlacement): TableOrigin {
    const expected = TableOrigin.expected();
    return new TableOrigin({
      headerRowIndex: SheetIndex.row(startRowIndex ?? expected.headerRowIndex),
      startColIndex: SheetIndex.col(startColumnIndex ?? expected.startColIndex),
    });
  },
};

function toGoogleTables(
  sheet: FakeSheetState,
  tables: readonly FakeTableState[],
): Table[] | undefined {
  if (tables.length === 0) return undefined;
  return tables.map((table) => ({
    tableId: table.tableId,
    ...(table.name !== undefined ? { name: table.name } : {}),
    range: {
      startRowIndex: table.startRowIndex,
      endRowIndex: table.endRowIndex,
      startColumnIndex: table.startColumnIndex,
      endColumnIndex: table.endColumnIndex,
    },
    columnProperties: columnProperties(sheet, table),
  }));
}

// An absent bound on the filter is unbounded, as in a Sheets GridRange.
function overlaps(table: FakeTableState, range: GridRange): boolean {
  return (
    table.startRowIndex < (range.endRowIndex ?? Infinity) &&
    (range.startRowIndex ?? 0) < table.endRowIndex &&
    table.startColumnIndex < (range.endColumnIndex ?? Infinity) &&
    (range.startColumnIndex ?? 0) < table.endColumnIndex
  );
}

function columnProperties(
  sheet: FakeSheetState,
  table: FakeTableState,
): TableColumnProperties[] | undefined {
  const {
    columnValidationValues = {},
    columnValidationConditionTypes = {},
    columnTypes = {},
    startColumnIndex,
    endColumnIndex,
  } = table;
  if (endColumnIndex <= startColumnIndex) {
    return undefined;
  }
  const headerRow = sheet.rows[table.startRowIndex] ?? [];
  return Array.from(
    { length: endColumnIndex - startColumnIndex },
    (_, tableRelativeIndex) => {
      const colIndex = startColumnIndex + tableRelativeIndex;
      const colProps: TableColumnProperties = {};
      if (tableRelativeIndex !== 0) {
        colProps.columnIndex = tableRelativeIndex;
      }
      const columnName = headerName(headerRow[colIndex]);
      if (columnName !== undefined) {
        colProps.columnName = columnName;
      }
      const columnType = columnTypes[colIndex];
      if (columnType) {
        colProps.columnType = columnType;
      }
      const values = columnValidationValues[colIndex];
      const conditionType = columnValidationConditionTypes[colIndex];
      if (values || conditionType) {
        colProps.dataValidationRule = {
          condition: {
            ...(conditionType ? { type: conditionType } : {}),
            ...(values
              ? {
                  values: values.map((userEnteredValue) => ({
                    userEnteredValue,
                  })),
                }
              : {}),
          },
        };
      }
      return colProps;
    },
  );
}

function headerName(cell: FakeCell | undefined): string | undefined {
  const value = cell === undefined ? null : fakeCells.value(cell);
  return typeof value === "string" && value !== "" ? value : undefined;
}
