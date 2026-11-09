import { Val } from "@byronbroughten/utils/val";

import type { CellValue } from "../../CellValues/cellValues";
import type { ConditionalFormatRule } from "../../RawSource/ConditionalFormat";
import type {
  GridCellSnapshot,
  GridRangeProps,
  SheetSnapshot,
  SpreadsheetSnapshot,
  TableColumnSnapshot,
  TableSnapshot,
} from "../../RawSource/RawSource";
import { SheetIndex } from "../../RawSource/SheetIndex";
import { googleColor } from "./googleColor";

type GoogleSpreadsheet = GoogleAppsScript.Sheets.Schema.Spreadsheet;
type GoogleSheet = GoogleAppsScript.Sheets.Schema.Sheet;
type GoogleCellData = GoogleAppsScript.Sheets.Schema.CellData;

export const googleGrid = {
  toSpreadsheetSnapshot(spreadsheet: GoogleSpreadsheet): SpreadsheetSnapshot {
    return spreadsheetSnapshot(spreadsheet, () => undefined);
  },
  // The rule reader is passed in because conditionalFormats.ts imports this file.
  toSpreadsheetSnapshotWithRules(
    spreadsheet: GoogleSpreadsheet,
    conditionalFormatsOf: (sheet: GoogleSheet) => ConditionalFormatRule[],
  ): SpreadsheetSnapshot {
    return spreadsheetSnapshot(spreadsheet, conditionalFormatsOf);
  },
  toGridRangeProps(
    range: GoogleAppsScript.Sheets.Schema.GridRange,
  ): GridRangeProps {
    // Google omits zero-valued fields, so a gid-0 sheet or column A arrives absent; observed: docs/sheets-behaviour.md#empty-lists-and-zero-valued-fields.
    return {
      sheetId: range.sheetId ?? 0,
      startRowIndex: SheetIndex.row(range.startRowIndex ?? 0),
      ...(range.endRowIndex !== undefined
        ? { endRowIndex: SheetIndex.row(range.endRowIndex) }
        : {}),
      startColumnIndex: SheetIndex.col(range.startColumnIndex ?? 0),
      ...(range.endColumnIndex !== undefined
        ? { endColumnIndex: SheetIndex.col(range.endColumnIndex) }
        : {}),
    };
  },
};

function spreadsheetSnapshot(
  spreadsheet: GoogleSpreadsheet,
  conditionalFormatsOf: (
    sheet: GoogleSheet,
  ) => ConditionalFormatRule[] | undefined,
): SpreadsheetSnapshot {
  return {
    timeZone: spreadsheet.properties?.timeZone ?? null,
    sheets: Val.assert(spreadsheet.sheets, "spreadsheet.sheets").map((sheet) =>
      toSheetSnapshot(sheet, conditionalFormatsOf(sheet)),
    ),
  };
}

function toSheetSnapshot(
  sheet: GoogleSheet,
  conditionalFormats: ConditionalFormatRule[] | undefined,
): SheetSnapshot {
  const properties = Val.assert(sheet.properties, "sheet.properties");
  return {
    conditionalFormats,
    sheetGid: Val.assert(properties.sheetId, "sheetId"),
    title: properties.title ?? null,
    rowCount: properties.gridProperties?.rowCount,
    columnCount: properties.gridProperties?.columnCount,
    // Google omits a false hidden and a zero frozen count, as it does other zero-valued fields.
    tabProperties: {
      color: googleColor.toTabColor(properties.tabColorStyle),
      hidden: properties.hidden ?? false,
      frozenRows: properties.gridProperties?.frozenRowCount ?? 0,
      frozenColumns: properties.gridProperties?.frozenColumnCount ?? 0,
    },
    tables: sheet.tables?.map(toTableSnapshot),
    gridBlocks: sheet.data?.map((block) => {
      const columnCount = (block.columnMetadata ?? []).length;
      return {
        startColumn: SheetIndex.col(block.startColumn ?? 0),
        startRow: SheetIndex.row(block.startRow ?? 0),
        columnCount,
        rows: (block.rowData ?? []).map((row) => ({
          cells: Array.from({ length: columnCount }, (_, colOffset) =>
            toGridCell(row.values?.[colOffset]),
          ),
        })),
      };
    }),
  };
}

function toTableSnapshot(
  table: NonNullable<GoogleSheet["tables"]>[number],
): TableSnapshot {
  const range = table.range;
  return {
    tableId: Val.assert(table.tableId, "tableId"),
    name: table.name ?? "",
    startRowIndex: SheetIndex.row(
      Val.assert(range?.startRowIndex, "startRowIndex"),
    ),
    endRowIndex: SheetIndex.row(Val.assert(range?.endRowIndex, "endRowIndex")),
    startColumnIndex: SheetIndex.col(
      Val.assert(range?.startColumnIndex, "startColumnIndex"),
    ),
    endColumnIndex: SheetIndex.col(
      Val.assert(range?.endColumnIndex, "endColumnIndex"),
    ),
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
    // The code assumes Google omits a zero columnIndex too; observed: docs/sheets-behaviour.md#empty-lists-and-zero-valued-fields.
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
    formula: cell.userEnteredValue?.formulaValue,
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
