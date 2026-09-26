import type { FakeCell } from "../fakeSheetsService";
import { fakeCells } from "./fakeCells";
import {
  type FakeSheetState,
  fakeSpreadsheet,
  type FakeTableState,
} from "./fakeSpreadsheet";

type Table = GoogleAppsScript.Sheets.Schema.Table;
type TableColumnProperties =
  GoogleAppsScript.Sheets.Schema.TableColumnProperties;

export const fakeTables = {
  // Google's shape for a sheet's Tables, as `get` returns them and the grid view shows them.
  googleTables(
    sheet: FakeSheetState,
    isFilteredFetch: boolean,
  ): Table[] | undefined {
    const tables = fakeSpreadsheet.tables(sheet);
    if (tables.length === 0) return undefined;
    if (sheet.isTableHiddenFromFilteredFetch && isFilteredFetch) {
      return undefined;
    }
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
  },
};

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
