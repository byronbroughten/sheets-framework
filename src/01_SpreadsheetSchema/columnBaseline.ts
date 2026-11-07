export interface ColumnValidationRule {
  condition: string;
  values: string[];
}
export interface ColumnTypeBaseline {
  // null: the column is left on Automatic.
  type: string | null;
  rule: ColumnValidationRule | null;
}
export type ActionRowCellBaseline =
  { kind: "checkbox" } | { kind: "groupHeading2"; text: string };
