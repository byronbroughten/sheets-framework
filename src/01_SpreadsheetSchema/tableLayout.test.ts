import { describe, expect, it } from "vitest";

import { headRows } from "./headRows";
import { tableLayout } from "./tableLayout";
import { uniformRows } from "./uniformRows";

const headRowOffsets = Object.values(tableLayout.headRowOffsets);

describe("tableLayout", () => {
  it("puts each head row a non-negative whole number of rows above the header", () => {
    headRowOffsets.forEach((offset) => {
      expect(Number.isInteger(offset)).toBe(true);
      expect(offset).toBeGreaterThanOrEqual(0);
    });
  });

  it("puts the header row itself at offset 0", () => {
    expect(tableLayout.headRowOffsets.header).toBe(0);
  });

  it("shares a row only between the action row and the level-2 group heading", () => {
    const { action, groupHeading2, ...rest } = tableLayout.headRowOffsets;
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
    expect(uniformRows.index("header")).toBe(-1);
    expect(uniformRows.index("action")).toBe(-2);
    expect(uniformRows.index("groupHeading1")).toBe(-3);
    expect(uniformRows.index("columnId")).toBe(-4);
  });
});

describe("headRows", () => {
  it("puts each head role's row above body row 0, the header just above it", () => {
    expect(headRows.index("header")).toBe(-1);
    expect(headRows.index("action")).toBe(-2);
    expect(headRows.index("groupHeading2")).toBe(-2);
    expect(headRows.index("groupHeading1")).toBe(-3);
    expect(headRows.index("columnId")).toBe(-4);
  });

  it("names every role a head row holds, so the shared row has two", () => {
    expect(headRows.rolesAt(-1)).toEqual(["header"]);
    expect(headRows.rolesAt(-2)).toEqual(["action", "groupHeading2"]);
    expect(headRows.rolesAt(-4)).toEqual(["columnId"]);
  });

  it("throws for an index no head row sits at", () => {
    expect(() => headRows.rolesAt(0)).toThrow(/not a head row/);
    expect(() => headRows.rolesAt(-5)).toThrow(/not a head row/);
  });
});
