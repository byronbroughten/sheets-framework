// What the name found, not what to do about it: the caller owns the wording.
export type RowIdByName =
  | { found: "one"; rowId: string; rowIndex: number }
  | { found: "none" }
  | { found: "many"; rowCount: number };
