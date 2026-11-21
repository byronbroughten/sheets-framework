import { a1Notation } from "./a1Notation";
import { fakeCells } from "./fakeCells";
import { fakeGrid } from "./fakeGrid";
import {
  type FakeSheetState,
  type FakeSpreadsheet,
  fakeSpreadsheet,
  type FakeTableState,
} from "./fakeSpreadsheet";
import { fieldMasks } from "./fieldMasks";
import { unreplayed } from "./unreplayed";

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
      throw unreplayed(
        "addTable columnProperties; send them in an updateTable",
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
  // Measured live: the cells above the header stay and rules wholly over the Table go; a rule reaching past it was never probed. An outside formula's `name[header]` becomes the A1 range of the column's old body rows (not #REF!); quoting a sheet title that isn't a plain identifier was never probed, and any other reference form throws.
  deleteTable(
    spreadsheet: FakeSpreadsheet,
    request: GoogleAppsScript.Sheets.Schema.DeleteTableRequest,
  ): Response {
    const { sheet, table } = fakeSpreadsheet.table(
      spreadsheet,
      request.tableId,
    );
    rewriteReferencesToRanges(spreadsheet, sheet, table);
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

function rewriteReferencesToRanges(
  spreadsheet: FakeSpreadsheet,
  sheet: FakeSheetState,
  table: FakeTableState,
): void {
  if (table.name === undefined) return;
  const sheetPrefix = /^[A-Za-z_]\w*$/.test(sheet.title)
    ? sheet.title
    : `'${sheet.title.replaceAll("'", "''")}'`;
  const rangesByHeader = new Map<string, string>();
  for (let col = table.startColumnIndex; col < table.endColumnIndex; col++) {
    const header = fakeCells.value(
      fakeGrid.cell(sheet, table.startRowIndex, col),
    );
    const letters = a1Notation.columnLetters(col);
    rangesByHeader.set(
      String(header),
      `${sheetPrefix}!$${letters}$${table.startRowIndex + 2}:$${letters}$${table.endRowIndex}`,
    );
  }
  const name = table.name.replaceAll(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const reference = new RegExp(`(?<![\\w.])${name}\\[([^\\]]*)\\]`, "g");
  const anyReference = new RegExp(`(?<![\\w.])${name}\\[`);
  for (const each of spreadsheet.sheets) {
    each.rows.forEach((row, rowIndex) => {
      row.forEach((cell, colIndex) => {
        if (typeof cell !== "object" || cell?.isFormula !== true) return;
        // The Table's own cells are blanked next, so what they read doesn't matter.
        if (each === sheet && isInTable(table, rowIndex, colIndex)) return;
        const rewritten = String(cell.value).replace(
          reference,
          (whole, header: string) => rangesByHeader.get(header) ?? whole,
        );
        if (anyReference.test(rewritten)) {
          throw unreplayed(
            `deleteTable for a formula reading Table "${table.name}" in a form other than name[header]: ${cell.value}`,
          );
        }
        if (rewritten === cell.value) return;
        fakeGrid.setCell(each, rowIndex, colIndex, {
          ...cell,
          value: rewritten,
          ...(cell.formula !== undefined ? { formula: rewritten } : {}),
        });
      });
    });
  }
}

function isInTable(
  table: FakeTableState,
  rowIndex: number,
  colIndex: number,
): boolean {
  return (
    rowIndex >= table.startRowIndex &&
    rowIndex < table.endRowIndex &&
    colIndex >= table.startColumnIndex &&
    colIndex < table.endColumnIndex
  );
}

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
  // Observed: docs/sheets-behaviour.md#creating-and-updating-tables.
  if (
    column.columnType === "DROPDOWN" &&
    column.dataValidationRule?.condition?.type === undefined
  ) {
    throw new Error("Invalid updateTable: No conditionType specified.");
  }
  const rule = column.dataValidationRule ?? {};
  const ruleFields = [
    ...Object.keys(rule).filter((field) => field !== "condition"),
    ...(rule.condition?.values ?? []).flatMap((value) =>
      Object.keys(value).filter((field) => field !== "userEnteredValue"),
    ),
  ];
  if (ruleFields.length > 0) {
    throw unreplayed(`updateTable dataValidationRule field "${ruleFields[0]}"`);
  }
}
