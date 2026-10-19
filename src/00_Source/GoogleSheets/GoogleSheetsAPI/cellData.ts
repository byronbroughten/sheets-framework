import { Obj } from "@byronbroughten/utils/obj";

import type { CellValue } from "../../CellValues/cellValues";
import type {
  FillCellOperation,
  FillColumnOperation,
} from "../../RawSource/RawSource";
import type { RgbColor } from "../../RawSource/RgbColor";
import type { ModeledRequest } from "../GoogleSheetsAPI";
import { googleColor } from "./googleColor";

type GoogleCellData = GoogleAppsScript.Sheets.Schema.CellData;
type UserEnteredValue = NonNullable<
  NonNullable<
    GoogleAppsScript.Sheets.Schema.UpdateCellsRequest["rows"]
  >[number]["values"]
>[number]["userEnteredValue"];

interface CellDataChange {
  value?: CellValue;
  backgroundColor?: RgbColor;
}

type GoogleGridRange = GoogleAppsScript.Sheets.Schema.GridRange;

interface CellRun {
  sheetId: number;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
}

export const cellDataRequests = {
  fillColumn(operation: FillColumnOperation): ModeledRequest[] {
    return formulaAndCellDataRequests(
      operation,
      {
        sheetId: operation.sheetId,
        rowIndex: operation.startRowIndex,
        columnIndex: operation.colIndex,
        rowCount: operation.endRowIndex - operation.startRowIndex,
      },
      (range, cell, fields) => ({ repeatCell: { range, cell, fields } }),
    );
  },
  fillCell(operation: FillCellOperation): ModeledRequest[] {
    return formulaAndCellDataRequests(
      operation,
      {
        sheetId: operation.sheetId,
        rowIndex: operation.rowIndex,
        columnIndex: operation.colIndex,
        rowCount: 1,
      },
      (range, cell, fields) => ({
        updateCells: { range, rows: [{ values: [cell] }], fields },
      }),
    );
  },
};

function formulaAndCellDataRequests(
  change: CellDataChange & { formula?: string },
  run: CellRun,
  cellRequest: (
    range: GoogleGridRange,
    cell: GoogleCellData,
    fields: string,
  ) => ModeledRequest,
): ModeledRequest[] {
  const requests: ModeledRequest[] = [];
  if (change.formula !== undefined) {
    requests.push(formulaPasteDataRequest(run, change.formula));
  }
  if (!cellChange.hasCellData(change)) return requests;
  requests.push(
    cellRequest(
      cellRunGridRange(run),
      cellChange.toCellData(change),
      cellChange.fieldMask(change),
    ),
  );
  return requests;
}

function cellRunGridRange(run: CellRun): GoogleGridRange {
  return {
    sheetId: run.sheetId,
    startRowIndex: run.rowIndex,
    endRowIndex: run.rowIndex + run.rowCount,
    startColumnIndex: run.columnIndex,
    endColumnIndex: run.columnIndex + 1,
  };
}

function formulaPasteDataRequest(run: CellRun, formula: string): ModeledRequest {
  const field = `"${formula.replaceAll('"', '""')}"`;
  return {
    pasteData: {
      coordinate: {
        sheetId: run.sheetId,
        rowIndex: run.rowIndex,
        columnIndex: run.columnIndex,
      },
      data: Array.from({ length: run.rowCount }, () => field).join("\n"),
      delimiter: "\t",
      type: "PASTE_FORMULA",
    },
  };
}

const cellChangeFields = {
  value: "userEnteredValue",
  backgroundColor: "userEnteredFormat.backgroundColor",
} as const satisfies Record<keyof CellDataChange, string>;

const cellChange = {
  hasCellData(change: CellDataChange): boolean {
    return change.value !== undefined || change.backgroundColor !== undefined;
  },
  fieldMask(change: CellDataChange): string {
    return Obj.keys(cellChangeFields)
      .filter((key) => change[key] !== undefined)
      .map((key) => cellChangeFields[key])
      .join(",");
  },
  toCellData(change: CellDataChange): GoogleCellData {
    const data: GoogleCellData = {};
    if (change.value !== undefined) {
      data.userEnteredValue = cellValueToUserEntered(change.value);
    }
    if (change.backgroundColor !== undefined) {
      data.userEnteredFormat = {
        backgroundColor: googleColor.fromRgb(change.backgroundColor),
      };
    }
    return data;
  },
};

function cellValueToUserEntered(value: CellValue): UserEnteredValue {
  if (typeof value === "string") {
    return { stringValue: value };
  }
  if (typeof value === "number") {
    return { numberValue: value };
  }
  if (typeof value === "boolean") {
    return { boolValue: value };
  }
  throw new Error(
    `Cannot make user entered value for unsupported type "${typeof value}".`,
  );
}
