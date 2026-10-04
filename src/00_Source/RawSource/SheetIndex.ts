declare const sheetRowIndex: unique symbol;
declare const sheetColIndex: unique symbol;

// Counted from the sheet's top-left cell, as Google counts; Raw counts from the Table.
export type SheetRowIndex = number & { readonly [sheetRowIndex]: true };
export type SheetColIndex = number & { readonly [sheetColIndex]: true };

export const SheetIndex = {
  isRow(value: number): value is SheetRowIndex {
    return Number.isInteger(value) && value >= 0;
  },
  isCol(value: number): value is SheetColIndex {
    return Number.isInteger(value) && value >= 0;
  },
  row(value: number): SheetRowIndex {
    if (SheetIndex.isRow(value)) return value;
    throw new Error(notAnIndexMessage(value, "row"));
  },
  col(value: number): SheetColIndex {
    if (SheetIndex.isCol(value)) return value;
    throw new Error(notAnIndexMessage(value, "column"));
  },
};

function notAnIndexMessage(value: number, dimension: "row" | "column"): string {
  return `${value} is not a sheet ${dimension} index: it must be a whole number of 0 or more.`;
}
