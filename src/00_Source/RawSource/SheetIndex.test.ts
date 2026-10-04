import { describe, expect, it } from "vitest";

import { assertNotType, type IsExactly } from "../../testSupport/typeAssertions";
import type { FillCellOperation, GridRangeProps } from "./RawSource";
import { type SheetColIndex, SheetIndex, type SheetRowIndex } from "./SheetIndex";

describe("SheetIndex", () => {
  it("brands a sheet row and column apart from a plain, Table-relative number", () => {
    assertNotType<IsExactly<SheetRowIndex, number>>(false);
    assertNotType<IsExactly<SheetColIndex, number>>(false);
    assertNotType<IsExactly<SheetRowIndex, SheetColIndex>>(false);
  });

  it("refuses a Table-relative index where Google expects a sheet coordinate", () => {
    const tableRowIndex = 0;
    const range: GridRangeProps = {
      sheetId: 1,
      // @ts-expect-error a Table-relative row index is not a sheet row index
      startRowIndex: tableRowIndex,
    };
    const fill: FillCellOperation = {
      kind: "fillCell",
      sheetId: 1,
      rowIndex: SheetIndex.row(4),
      // @ts-expect-error a sheet row index is not a sheet column index
      colIndex: SheetIndex.row(0),
    };
    expect([range.sheetId, fill.rowIndex]).toEqual([1, 4]);
  });

  it("accepts a whole number of 0 or more", () => {
    expect(SheetIndex.row(0)).toBe(0);
    expect(SheetIndex.col(27)).toBe(27);
  });

  it("refuses a negative or fractional index", () => {
    expect(() => SheetIndex.row(-1)).toThrow("-1 is not a sheet row index");
    expect(() => SheetIndex.col(1.5)).toThrow("1.5 is not a sheet column index");
  });
});
