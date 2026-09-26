import { fakeGrid } from "./fakeGrid";
import {
  type FakeSheetState,
  type FakeSpreadsheet,
  fakeSpreadsheet,
  type FakeTableState,
} from "./fakeSpreadsheet";

type Response = GoogleAppsScript.Sheets.Schema.Response;
type Table = GoogleAppsScript.Sheets.Schema.Table;
type TableColumnProperties =
  GoogleAppsScript.Sheets.Schema.TableColumnProperties;

export const tableReplays = {
  // Measured live: addTable's own column names land three columns right, so the adapter never sends them.
  addTable(
    spreadsheet: FakeSpreadsheet,
    request: GoogleAppsScript.Sheets.Schema.AddTableRequest,
  ): Response {
    const table = request.table ?? {};
    if (table.columnProperties !== undefined) {
      throw new Error(
        "The fake Sheets service does not replay addTable columnProperties; send them in an updateTable.",
      );
    }
    const range = table.range ?? {};
    const sheet = fakeSpreadsheet.sheet(spreadsheet, range.sheetId);
    const bounded = fakeGrid.boundedRange(sheet, range);
    const tableState: FakeTableState = {
      tableId:
        table.tableId ??
        `fake-table-${sheet.sheetId}-added-${fakeSpreadsheet.tables(sheet).length}`,
      ...(table.name !== undefined ? { name: table.name } : {}),
      ...bounded,
    };
    if (sheet.table === undefined) {
      sheet.table = tableState;
    } else {
      sheet.extraTables.push(tableState);
    }
    return {};
  },
  // Measured live: a field mask replaces only its fields.
  updateTable(
    spreadsheet: FakeSpreadsheet,
    request: GoogleAppsScript.Sheets.Schema.UpdateTableRequest,
  ): Response {
    const update = request.table ?? {};
    const { sheet, table } = fakeSpreadsheet.table(spreadsheet, update.tableId);
    (request.fields ?? "").split(",").forEach((rawField) => {
      const field = rawField.trim();
      if (field === "name") {
        table.name = update.name;
      } else if (field === "range") {
        Object.assign(table, fakeGrid.boundedRange(sheet, update.range ?? {}));
      } else if (field === "columnProperties") {
        replaceColumnProperties(sheet, table, update);
      } else {
        throw new Error(
          `The fake Sheets service does not replay updateTable field "${field}".`,
        );
      }
    });
    return {};
  },
};

// Measured live: the sent list replaces the Table's, so an unsent column keeps its header but loses its type.
function replaceColumnProperties(
  sheet: FakeSheetState,
  table: FakeTableState,
  update: Table,
): void {
  const columnProperties = update.columnProperties ?? [];
  columnProperties.forEach(validateColumnName);
  table.columnTypes = {};
  table.columnValidationValues = {};
  table.columnValidationConditionTypes = {};
  const {
    columnTypes,
    columnValidationValues,
    columnValidationConditionTypes,
  } = table;
  columnProperties.forEach((column) => {
    const colIndex = table.startColumnIndex + (column.columnIndex ?? 0);
    fakeGrid.setCell(
      sheet,
      table.startRowIndex,
      colIndex,
      column.columnName ?? null,
    );
    if (column.columnType !== undefined) {
      columnTypes[colIndex] = column.columnType;
    }
    const condition = column.dataValidationRule?.condition;
    if (condition?.type !== undefined) {
      columnValidationConditionTypes[colIndex] = condition.type;
    }
    const values = (condition?.values ?? []).flatMap((value) =>
      value.userEnteredValue === undefined ? [] : [value.userEnteredValue],
    );
    if (values.length > 0) columnValidationValues[colIndex] = values;
  });
}

function validateColumnName(column: TableColumnProperties): void {
  if (column.columnName === undefined || column.columnName === "") {
    throw new Error(
      `Invalid updateTable: column ${column.columnIndex ?? 0} has no columnName.`,
    );
  }
}
