import { describe, expect, it } from "vitest";

import { idPrefixes } from "./idPrefixes";

describe("idPrefixes.fromTitle", () => {
  it("abbreviates a one-word title to the first letter plus consonants, up to 3", () => {
    expect(idPrefixes.fromTitle("Product", new Set())).toBe("prd");
    expect(idPrefixes.fromTitle("Item", new Set())).toBe("itm");
    expect(idPrefixes.fromTitle("Schedule", new Set())).toBe("sch");
  });

  it("takes the first letter of each word and tops up from the last word's consonants to 3", () => {
    expect(idPrefixes.fromTitle("Run Item", new Set())).toBe("rit");
  });

  it("keeps a title shorter than 3 consonants", () => {
    expect(idPrefixes.fromTitle("Id", new Set())).toBe("id");
  });

  it("steps up with the next consonant on a collision, then a number suffix", () => {
    expect(idPrefixes.fromTitle("Product", new Set(["prd"]))).toBe("prdc");
    expect(
      idPrefixes.fromTitle("Product", new Set(["prd", "prdc", "prdct"])),
    ).toBe("prd2");
    expect(
      idPrefixes.fromTitle(
        "Product",
        new Set(["prd", "prdc", "prdct", "prd2"]),
      ),
    ).toBe("prd3");
  });

  it("uses s plus a number suffix when the title has no letters", () => {
    expect(idPrefixes.fromTitle("2024", new Set())).toBe("s");
    expect(idPrefixes.fromTitle("2024", new Set(["s"]))).toBe("s2");
  });

  it("drops punctuation and digits before abbreviating", () => {
    expect(idPrefixes.fromTitle("Item-2B!", new Set())).toBe("itm");
  });

  it("returns only lowercase letters and digits", () => {
    const prefixes = [
      idPrefixes.fromTitle("Product", new Set()),
      idPrefixes.fromTitle("Run Item", new Set()),
      idPrefixes.fromTitle("2024", new Set(["s"])),
      idPrefixes.fromTitle("Item-2B!", new Set()),
      idPrefixes.fromTitle("Product", new Set(["prd", "prdc", "prdct"])),
    ];
    prefixes.forEach((prefix) => {
      expect(prefix).toMatch(/^[a-z0-9]+$/);
    });
  });
});
