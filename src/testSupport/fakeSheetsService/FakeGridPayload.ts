import type {
  FakeCell,
  FakeCellValue,
  FakeRichCellValue,
} from "../fakeSheetsService";
import type { FakeSheetState } from "./fakeSpreadsheet";

type GoogleCellData = GoogleAppsScript.Sheets.Schema.CellData;
type GoogleSheetData = GoogleAppsScript.Sheets.Schema.Sheet["data"];

// A fixture sheet's rows as a grid read's data blocks, shaped by the read's field mask.
export class FakeGridPayload {
  private isFormulaOnlyMask: boolean;
  constructor(fields: string | undefined) {
    this.isFormulaOnlyMask =
      fields?.includes("userEnteredValue(formulaValue)") ?? false;
  }
  sheetData({
    rows,
    rowsWithNoGridData = [],
    rowsWithNoGridBlock = [],
  }: FakeSheetState): GoogleSheetData | undefined {
    if (rows.length === 0) {
      return undefined;
    }
    const columnCount = Math.max(0, ...rows.map((row) => row.length));
    const noGridDataRows = new Set(rowsWithNoGridData);
    const noGridBlockRows = new Set(rowsWithNoGridBlock);

    // Real rowData blocks run contiguously from a real startRow, so either
    // kind of absent row ends the block it was in.
    const blocks: NonNullable<GoogleSheetData> = [];
    let currentBlockRows: GoogleAppsScript.Sheets.Schema.RowData[] = [];
    let currentBlockStart: number | null = null;
    function columnMetadata(): GoogleAppsScript.Sheets.Schema.DimensionProperties[] {
      return Array.from({ length: columnCount }, () => ({}));
    }
    function flushCurrentBlock(): void {
      if (currentBlockStart !== null) {
        blocks.push({
          startColumn: 0,
          startRow: currentBlockStart,
          columnMetadata: columnMetadata(),
          rowData: currentBlockRows,
        });
      }
      currentBlockRows = [];
      currentBlockStart = null;
    }
    rows.forEach((row, rowIndex) => {
      if (noGridBlockRows.has(rowIndex)) {
        flushCurrentBlock();
        return;
      }
      if (noGridDataRows.has(rowIndex)) {
        flushCurrentBlock();
        // Every grid column described, no rowData — the measured real shape.
        blocks.push({
          startColumn: 0,
          startRow: rowIndex,
          columnMetadata: columnMetadata(),
        });
        return;
      }
      currentBlockStart ??= rowIndex;
      currentBlockRows.push({
        values: Array.from(
          { length: columnCount },
          (_, colIndex): GoogleCellData =>
            this._cellData(row[colIndex] ?? null),
        ),
      });
    });
    flushCurrentBlock();
    return blocks;
  }
  // Like Google: under userEnteredValue(formulaValue), a non-formula value comes back as {}; observed: docs/sheets-behaviour.md#reading-formula-text-beside-effective-values.
  private _cellData(cell: FakeCell): GoogleCellData {
    if (cell === null) {
      return {};
    }
    const rich: FakeRichCellValue =
      typeof cell === "object" ? cell : { value: cell };
    const data: GoogleCellData = {
      effectiveValue: fakeValueToExtendedValue(rich.value),
    };
    if (rich.isFormula) {
      data.userEnteredValue = {
        formulaValue: rich.formula ?? formulaText(rich.value),
      };
    } else if (this.isFormulaOnlyMask && rich.value !== null) {
      data.userEnteredValue = {};
    }
    if (rich.numberFormatType) {
      data.effectiveFormat = {
        numberFormat: { type: rich.numberFormatType },
      };
    }
    if (rich.dataValidationConditionType) {
      data.dataValidation = {
        condition: { type: rich.dataValidationConditionType },
      };
    }
    return data;
  }
}

function fakeValueToExtendedValue(
  value: FakeCellValue,
): GoogleAppsScript.Sheets.Schema.ExtendedValue | undefined {
  if (value === null) {
    return undefined;
  }
  if (typeof value === "string") {
    return { stringValue: value };
  }
  if (typeof value === "number") {
    return { numberValue: value };
  }
  return { boolValue: value };
}

// A replayed formula keeps its text; a fixture's formula cell gives only its computed value.
function formulaText(value: FakeCellValue): string {
  return typeof value === "string" && value.startsWith("=")
    ? value
    : "=FAKE_FORMULA()";
}
