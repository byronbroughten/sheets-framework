import { fakeGrid } from "./fakeGrid";
import {
  type FakeSheetState,
  type FakeSpreadsheet,
  fakeSpreadsheet,
  type FakeTableState,
} from "./fakeSpreadsheet";
import { fieldMasks } from "./fieldMasks";

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
        `fake-table-${sheet.sheetId}-added-${sheet.tables.length}`,
      ...(table.name !== undefined ? { name: table.name } : {}),
      ...bounded,
    };
    sheet.tables.push(tableState);
    return {};
  },
  // Measured live: the cells above the header stay and rules wholly over the Table go; a rule reaching past it was never probed.
  deleteTable(
    spreadsheet: FakeSpreadsheet,
    request: GoogleAppsScript.Sheets.Schema.DeleteTableRequest,
  ): Response {
    const { sheet, table } = fakeSpreadsheet.table(
      spreadsheet,
      request.tableId,
    );
    sheet.tables = sheet.tables.filter((each) => each !== table);
    fakeGrid.forEachCell(table, (rowIndex, colIndex) => {
      fakeGrid.setCell(sheet, rowIndex, colIndex, null);
    });
    if (sheet.conditionalFormats !== undefined) {
      sheet.conditionalFormats = sheet.conditionalFormats.filter(
        (rule) => !isWhollyOver(rule, table),
      );
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
    fieldMasks.replay(request.fields, "updateTable", {
      name() {
        table.name = update.name;
      },
      range() {
        Object.assign(table, fakeGrid.boundedRange(sheet, update.range ?? {}));
      },
      columnProperties() {
        replaceColumnProperties(sheet, table, update);
      },
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
  columnProperties.forEach(validateColumn);
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

// An open bound reaches the sheet's edge, so a rule with one is never wholly over a Table.
function isWhollyOver(
  rule: GoogleAppsScript.Sheets.Schema.ConditionalFormatRule,
  table: FakeTableState,
): boolean {
  return (rule.ranges ?? []).every(
    (range) =>
      range.startRowIndex !== undefined &&
      range.startColumnIndex !== undefined &&
      range.endRowIndex !== undefined &&
      range.endColumnIndex !== undefined &&
      table.startRowIndex <= range.startRowIndex &&
      range.endRowIndex <= table.endRowIndex &&
      table.startColumnIndex <= range.startColumnIndex &&
      range.endColumnIndex <= table.endColumnIndex,
  );
}

function validateColumn(column: TableColumnProperties): void {
  if (column.columnName === undefined || column.columnName === "") {
    throw new Error(
      `Invalid updateTable: column ${column.columnIndex ?? 0} has no columnName.`,
    );
  }
  const rule = column.dataValidationRule ?? {};
  const ruleFields = [
    ...Object.keys(rule).filter((field) => field !== "condition"),
    ...(rule.condition?.values ?? []).flatMap((value) =>
      Object.keys(value).filter((field) => field !== "userEnteredValue"),
    ),
  ];
  if (ruleFields.length > 0) {
    throw new Error(
      `The fake Sheets service does not replay updateTable dataValidationRule field "${ruleFields[0]}".`,
    );
  }
}
