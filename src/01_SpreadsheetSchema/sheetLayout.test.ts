import { describe, expect, it } from "vitest";

import { sheetLayout } from "./sheetLayout";

const uniformRowIndexes = [
  sheetLayout.colIdRowIndex,
  sheetLayout.colGroupHeadingRowIndex,
  sheetLayout.actionRowIndex,
  sheetLayout.tableHeaderRowIndex,
];

describe("sheetLayout", () => {
  it("puts each uniform row on a non-negative integer index", () => {
    uniformRowIndexes.forEach((index) => {
      expect(Number.isInteger(index)).toBe(true);
      expect(index).toBeGreaterThanOrEqual(0);
    });
  });

  it("gives no two uniform rows the same row", () => {
    expect(new Set(uniformRowIndexes).size).toBe(uniformRowIndexes.length);
  });

  it("puts no uniform row on the first data row", () => {
    expect(uniformRowIndexes).not.toContain(
      sheetLayout.tableHeaderRowIndex + 1,
    );
  });

  it("starts the Table on a non-negative integer column", () => {
    expect(Number.isInteger(sheetLayout.startTableColIndex)).toBe(true);
    expect(sheetLayout.startTableColIndex).toBeGreaterThanOrEqual(0);
  });
});
