import { describe, expect, it } from "vitest";

import { dimensionIds } from "./dimensionIds";

describe("dimensionIds.col / dimensionIds.row", () => {
  it("builds a column ID and a row ID with a given suffix", () => {
    expect(dimensionIds.col("hh", "name")).toBe("c:hh:name");
    expect(dimensionIds.row("hh", "1")).toBe("r:hh:1");
  });

  it("builds IDs with a random 7-character suffix when none is given", () => {
    const colId = dimensionIds.col("hh");
    expect(colId).toMatch(/^c:hh:[0-9a-zA-Z_-]{7}$/);
    expect(dimensionIds.row("hh")).toMatch(/^r:hh:[0-9a-zA-Z_-]{7}$/);
    expect(dimensionIds.col("hh")).not.toBe(colId);
  });

  it("throws when the ID prefix is empty", () => {
    expect(() => dimensionIds.col("")).toThrow();
    expect(() => dimensionIds.row("", "1")).toThrow();
  });
});

describe("dimensionIds.colIdPrefixOrUndefined", () => {
  it("recovers the ID prefix of a built column ID", () => {
    expect(dimensionIds.colIdPrefixOrUndefined(dimensionIds.col("hh"))).toBe(
      "hh",
    );
  });

  it("is undefined for a row ID", () => {
    expect(dimensionIds.colIdPrefixOrUndefined("r:hh:abc")).toBeUndefined();
  });

  it("is undefined for the wrong number of parts", () => {
    expect(dimensionIds.colIdPrefixOrUndefined("c:hh")).toBeUndefined();
    expect(dimensionIds.colIdPrefixOrUndefined("c:hh:abc:x")).toBeUndefined();
    expect(dimensionIds.colIdPrefixOrUndefined("c|hh|abc")).toBeUndefined();
  });

  it("is undefined for an empty part", () => {
    expect(dimensionIds.colIdPrefixOrUndefined("c::abc")).toBeUndefined();
    expect(dimensionIds.colIdPrefixOrUndefined("c:hh:")).toBeUndefined();
  });
});
