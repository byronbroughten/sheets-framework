import type { TableColumnValidationRule } from "../00_Source/RawSource/RawSource.ts";

export type ColumnValidationRule = TableColumnValidationRule;
export interface ColumnTypeBaseline {
  // null: the column is left on Automatic.
  type: string | null;
  rule: ColumnValidationRule | null;
}
export type ActionRowCellBaseline =
  { kind: "checkbox" } | { kind: "groupHeading2"; text: string };
