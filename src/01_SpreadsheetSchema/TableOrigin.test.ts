import { describe, expect, it } from "vitest";

import { SheetIndex } from "../00_Source/RawSource/SheetIndex";
import { TableOrigin } from "./TableOrigin";

describe("TableOrigin", () => {
  const lowerOrigin = new TableOrigin({
    headerRowIndex: SheetIndex.row(9),
    startColIndex: SheetIndex.col(2),
  });

  it("puts body row 0 on the sheet row just below the header", () => {
    expect(lowerOrigin.sheetRowIndex(0)).toBe(10);
    expect(lowerOrigin.rowIndex(SheetIndex.row(10))).toBe(0);
  });

  it("puts a head row at its offset above the header", () => {
    expect(lowerOrigin.sheetRowIndex(-1)).toBe(9);
    expect(lowerOrigin.sheetRowIndex(-4)).toBe(6);
  });

  it("puts the topmost head row the head-row count above the header", () => {
    expect(lowerOrigin.topHeadSheetRowIndex).toBe(6);
  });

  it("puts column 0 on the Table's first sheet column", () => {
    expect(lowerOrigin.sheetColIndex(0)).toBe(2);
    expect(lowerOrigin.colIndex(SheetIndex.col(4))).toBe(2);
  });

  it("numbers a row as the sheet's 1-based gutter shows it", () => {
    expect(lowerOrigin.rowNumber(0)).toBe(11);
    expect(lowerOrigin.rowNumber(-1)).toBe(10);
  });

  it("refuses an index that would land above the sheet's first row", () => {
    expect(() => TableOrigin.expected().sheetRowIndex(-5)).toThrow(
      "-1 is not a sheet row index",
    );
  });

  it("expects the head rows to start at the sheet's first cell", () => {
    const expected = TableOrigin.expected();
    expect(expected.sheetRowIndex(-4)).toBe(0);
    expect(expected.headerRowIndex).toBe(3);
    expect(expected.startColIndex).toBe(0);
  });
});
