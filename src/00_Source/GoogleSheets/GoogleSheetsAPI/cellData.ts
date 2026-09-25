import { Obj } from "../../../utils/Obj";
import type { CellValue } from "../../CellValues/cellValues";
import type {
  FillOperation,
  UpdateCellOperation,
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

export const cellDataRequests = {
  fill(operation: FillOperation): ModeledRequest[] {
    return formulaAndCellDataRequests(
      operation,
      {
        sheetId: operation.sheetId,
        rowIndex: operation.startRowIndex,
        columnIndex: operation.colIndex,
        rowCount: operation.endRowIndex - operation.startRowIndex,
      },
      (cell, fields) => ({
        repeatCell: {
          range: {
            sheetId: operation.sheetId,
            startRowIndex: operation.startRowIndex,
            endRowIndex: operation.endRowIndex,
            startColumnIndex: operation.colIndex,
            endColumnIndex: operation.colIndex + 1,
          },
          cell,
          fields,
        },
      }),
    );
  },
  updateCell(operation: UpdateCellOperation): ModeledRequest[] {
    return formulaAndCellDataRequests(
      operation,
      {
        sheetId: operation.sheetId,
        rowIndex: operation.rowIndex,
        columnIndex: operation.colIndex,
        rowCount: 1,
      },
      (cell, fields) => ({
        updateCells: {
          range: {
            sheetId: operation.sheetId,
            startRowIndex: operation.rowIndex,
            endRowIndex: operation.rowIndex + 1,
            startColumnIndex: operation.colIndex,
            endColumnIndex: operation.colIndex + 1,
          },
          rows: [{ values: [cell] }],
          fields,
        },
      }),
    );
  },
};

function formulaAndCellDataRequests(
  change: CellDataChange & { formula?: string },
  pasteProps: {
    sheetId: number;
    rowIndex: number;
    columnIndex: number;
    rowCount: number;
  },
  cellRequest: (cell: GoogleCellData, fields: string) => ModeledRequest,
): ModeledRequest[] {
  const requests: ModeledRequest[] = [];
  if (change.formula !== undefined) {
    requests.push(
      formulaPasteDataRequest({ ...pasteProps, formula: change.formula }),
    );
  }
  if (!cellChange.hasCellData(change)) return requests;
  requests.push(
    cellRequest(cellChange.toCellData(change), cellChange.fieldMask(change)),
  );
  return requests;
}

function formulaPasteDataRequest(props: {
  sheetId: number;
  rowIndex: number;
  columnIndex: number;
  formula: string;
  rowCount: number;
}): ModeledRequest {
  const field = `"${props.formula.replaceAll('"', '""')}"`;
  return {
    pasteData: {
      coordinate: {
        sheetId: props.sheetId,
        rowIndex: props.rowIndex,
        columnIndex: props.columnIndex,
      },
      data: Array.from({ length: props.rowCount }, () => field).join("\n"),
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
