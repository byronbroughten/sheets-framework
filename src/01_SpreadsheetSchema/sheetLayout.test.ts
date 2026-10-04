import { describe, expect, it } from "vitest";

import { sheetLayout } from "./sheetLayout";
import { uniformRows } from "./uniformRows";

const headRowOffsets = Object.values(sheetLayout.headRowOffsets);

describe("sheetLayout", () => {
  it("puts each head row a non-negative whole number of rows above the header", () => {
    headRowOffsets.forEach((offset) => {
      expect(Number.isInteger(offset)).toBe(true);
      expect(offset).toBeGreaterThanOrEqual(0);
    });
  });

  it("puts the header row itself at offset 0", () => {
    expect(sheetLayout.headRowOffsets.header).toBe(0);
  });

  it("shares a row only between the action row and the level-2 group heading", () => {
    const { action, groupHeading2, ...rest } = sheetLayout.headRowOffsets;
    expect(action).toBe(groupHeading2);
    const unshared = [action, ...Object.values(rest)];
    expect(new Set(unshared).size).toBe(unshared.length);
  });
});

describe("uniformRows", () => {
  it("puts every uniform row above body row 0", () => {
    Object.values(uniformRows.indexes()).forEach((rowIndex) => {
      expect(rowIndex).toBeLessThan(0);
    });
  });

  it("puts the header row just above body row 0 and the column ID row furthest up", () => {
    expect(uniformRows.index("tableHeader")).toBe(-1);
    expect(uniformRows.index("action")).toBe(-2);
    expect(uniformRows.index("colGroupName")).toBe(-3);
    expect(uniformRows.index("columnId")).toBe(-4);
  });
});
