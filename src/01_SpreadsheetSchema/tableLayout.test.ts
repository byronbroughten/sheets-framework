import { describe, expect, it } from "vitest";

import { headRows } from "./headRows";
import { tableLayout } from "./tableLayout";

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

  it("lists each head row's index once, a shared row included", () => {
    expect(headRows.indexes()).toEqual([-1, -2, -3, -4]);
    expect(headRows.isIndex(-2)).toBe(true);
    expect(headRows.isIndex(0)).toBe(false);
  });

  it("takes the head rows' extent from the largest offset, not a named role", () => {
    expect(headRows.countAboveHeader).toBe(3);
    expect(headRows.topIndex).toBe(headRows.index("columnId"));
    expect(headRows.lastAboveHeaderIndex).toBe(-2);
  });

  it("throws for an index no head row sits at", () => {
    expect(() => headRows.rolesAt(0)).toThrow(/not a head row/);
    expect(() => headRows.rolesAt(-5)).toThrow(/not a head row/);
  });
});
