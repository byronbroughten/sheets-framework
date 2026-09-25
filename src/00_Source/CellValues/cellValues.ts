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

const uniformRowValueNames = {
  tableHeader: "string",
  action: "boolean", // Should perhaps be "boolean" | "string"
  columnId: "string",
  colGroupName: "string",
} as const;

type UniformRowValueNames = typeof uniformRowValueNames;
export type UniformRowName = keyof UniformRowValueNames;
export type UniformRowValueName<UN extends UniformRowName> =
  UniformRowValueNames[UN];

export type UniformRowValue<UN extends UniformRowName> = CellValue<
  UniformRowValueName<UN>
>;

export function getUniformRowValueName<UN extends UniformRowName>(
  name: UN,
): UniformRowValueName<UN> {
  return uniformRowValueNames[name];
}
