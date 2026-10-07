import { Obj } from "../../utils/Obj";

const cellValues = {
  string: "" as string,
  number: 0 as number,
  date: 0 as number,
  boolean: false as boolean,
};

export const cellValueNames: CellValueName[] = Obj.keys(cellValues);
export type CellValueNameToValue = typeof cellValues;
export type CellValueName = keyof CellValueNameToValue;
export type CellValue<VN extends CellValueName = CellValueName> =
  CellValueNameToValue[VN];

// The blank removed, whether or not this value type ever had one.
export type NotEmpty<VL> = Exclude<VL, "">;

export const codebaseNameDelimiter = "_";
export type CodebaseNameDelimiter = typeof codebaseNameDelimiter;

