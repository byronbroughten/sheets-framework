import type { CellAfter, GoogleCellData } from "../flushClassifierTypes";

export const cellAfter = {
  written(cell: GoogleCellData | undefined): CellAfter {
    const value = cell?.userEnteredValue;
    if (value === undefined || Object.keys(value).length === 0) {
      return { kind: "cleared" };
    }
    return { kind: "value", value };
  },
  // Google parses pasted text as typed into a cell: a leading "=" is a formula, a numeral a number.
  pasted(field: string): CellAfter {
    if (field === "") return { kind: "cleared" };
    if (field.startsWith("=")) {
      return { kind: "value", value: { formulaValue: field } };
    }
    if (/^-?\d+(\.\d+)?$/.test(field)) {
      return { kind: "value", value: { numberValue: Number(field) } };
    }
    const upper = field.toUpperCase();
    if (upper === "TRUE" || upper === "FALSE") {
      return { kind: "value", value: { boolValue: upper === "TRUE" } };
    }
    return { kind: "value", value: { stringValue: field } };
  },
};
