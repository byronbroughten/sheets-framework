import { Val } from "../../../utils/Val";
import type { CellValue } from "../../CellValues/cellValues";
import type {
  GridCellSnapshot,
  GridRangeProps,
  SheetSnapshot,
  SpreadsheetSnapshot,
  TableColumnSnapshot,
  TableSnapshot,
} from "../../RawSource/RawSource";
import { googleColor } from "./googleColor";

type GoogleSpreadsheet = GoogleAppsScript.Sheets.Schema.Spreadsheet;
type GoogleSheet = GoogleAppsScript.Sheets.Schema.Sheet;
type GoogleCellData = GoogleAppsScript.Sheets.Schema.CellData;

export const googleGrid = {
  toSpreadsheetSnapshot(spreadsheet: GoogleSpreadsheet): SpreadsheetSnapshot {
    return {
      timeZone: spreadsheet.properties?.timeZone ?? null,
      sheets: Val.assert(spreadsheet.sheets, "spreadsheet.sheets").map(
        toSheetSnapshot,
      ),
    };
  },
  toGridRangeProps(
    range: GoogleAppsScript.Sheets.Schema.GridRange,
  ): GridRangeProps {
    // Google omits zero-valued fields, so a gid-0 sheet or column A arrives absent.
    return {
      sheetId: range.sheetId ?? 0,
      startRowIndex: range.startRowIndex ?? 0,
      ...(range.endRowIndex !== undefined
        ? { endRowIndex: range.endRowIndex }
        : {}),
      startColumnIndex: range.startColumnIndex ?? 0,
      ...(range.endColumnIndex !== undefined
        ? { endColumnIndex: range.endColumnIndex }
        : {}),
    };
  },
};

function toSheetSnapshot(sheet: GoogleSheet): SheetSnapshot {
  const properties = Val.assert(sheet.properties, "sheet.properties");
  return {
    sheetGid: Val.assert(properties.sheetId, "sheetId"),
    title: properties.title ?? null,
    tables: sheet.tables?.map(toTableSnapshot),
    gridBlocks: sheet.data?.map((block) => ({
      startColumn: block.startColumn ?? 0,
      startRow: block.startRow ?? 0,
      columnCount: (block.columnMetadata || []).length,
      rows: (block.rowData || []).map((row) => ({
        cells: Array.from(
          { length: (block.columnMetadata || []).length },
          (_, colOffset) => toGridCell(row.values?.[colOffset]),
        ),
      })),
    })),
  };
}

function toTableSnapshot(
  table: NonNullable<GoogleSheet["tables"]>[number],
): TableSnapshot {
  const range = table.range;
  return {
    tableId: Val.assert(table.tableId, "tableId"),
    name: table.name ?? "",
    startRowIndex: Val.assert(range?.startRowIndex, "startRowIndex"),
    endRowIndex: Val.assert(range?.endRowIndex, "endRowIndex"),
    startColumnIndex: Val.assert(range?.startColumnIndex, "startColumnIndex"),
    endColumnIndex: Val.assert(range?.endColumnIndex, "endColumnIndex"),
    columnProperties: (table.columnProperties ?? []).map(toTableColumnSnapshot),
  };
}

function toTableColumnSnapshot(
  colProps: NonNullable<
    NonNullable<GoogleSheet["tables"]>[number]["columnProperties"]
  >[number],
): TableColumnSnapshot {
  const values = (colProps.dataValidationRule?.condition?.values ?? [])
    .map((conditionValue) => conditionValue.userEnteredValue)
    .filter((value): value is string => value !== undefined);
  return {
    // The API omits columnIndex when it's zero.
    columnIndex: colProps.columnIndex ?? 0,
    ...(colProps.columnName !== undefined
      ? { columnName: colProps.columnName }
      : {}),
    columnType: colProps.columnType,
    dataValidationValues: values,
    dataValidationConditionType: colProps.dataValidationRule?.condition?.type,
  };
}

function toGridCell(
  cell: GoogleCellData | undefined,
): GridCellSnapshot | undefined {
  if (cell === undefined) return undefined;
  const backgroundColor = cell.userEnteredFormat?.backgroundColor;
  return {
    value: effectiveCellValue(cell),
    isFormula: cell.userEnteredValue?.formulaValue !== undefined,
    numberFormatType: cell.effectiveFormat?.numberFormat?.type,
    dataValidationConditionType: cell.dataValidation?.condition?.type,
    ...(backgroundColor !== undefined
      ? { backgroundColor: googleColor.toRgb(backgroundColor) }
      : {}),
  };
}

function effectiveCellValue(cell: GoogleCellData): CellValue | "" {
  const effectiveValue = cell.effectiveValue;
  if (effectiveValue === undefined) {
    return "";
  }
  if ("stringValue" in effectiveValue) {
    return Val.assert(effectiveValue.stringValue, "stringValue");
  }
  if ("boolValue" in effectiveValue) {
    return Val.assert(effectiveValue.boolValue, "boolValue");
  }
  if ("numberValue" in effectiveValue) {
    return Val.assert(effectiveValue.numberValue, "numberValue");
  }
  return "";
}
