import { describe, expect, it } from "vitest";

import { idPrefixes } from "./idPrefixes";

describe("idPrefixes.fromTitle", () => {
  it("abbreviates a one-word title to the first letter plus consonants, up to 3", () => {
    expect(
      idPrefixes.fromTitle({ title: "Product", prefixesInUse: new Set() }),
    ).toBe("prd");
    expect(
      idPrefixes.fromTitle({ title: "Item", prefixesInUse: new Set() }),
    ).toBe("itm");
    expect(
      idPrefixes.fromTitle({ title: "Schedule", prefixesInUse: new Set() }),
    ).toBe("sch");
  });

  it("takes the first letter of each word and tops up from the last word's consonants to 3", () => {
    expect(
      idPrefixes.fromTitle({ title: "Run Item", prefixesInUse: new Set() }),
    ).toBe("rit");
  });

  it("gives one prefix for sentence, title, snake and camel spellings", () => {
    [
      "add occ charge",
      "Add Occ Charge",
      "add_occ_charge",
      "addOccCharge",
    ].forEach((spelling) => {
      expect(
        idPrefixes.fromTitle({ title: spelling, prefixesInUse: new Set() }),
      ).toBe("aoc");
    });
  });

  it("keeps a title shorter than 3 consonants", () => {
    expect(
      idPrefixes.fromTitle({ title: "Id", prefixesInUse: new Set() }),
    ).toBe("id");
  });

  it("steps up with the next consonant on a collision, then a number suffix", () => {
    expect(
      idPrefixes.fromTitle({
        title: "Product",
        prefixesInUse: new Set(["prd"]),
      }),
    ).toBe("prdc");
    expect(
      idPrefixes.fromTitle({
        title: "Product",
        prefixesInUse: new Set(["prd", "prdc", "prdct"]),
      }),
    ).toBe("prd2");
    expect(
      idPrefixes.fromTitle({
        title: "Product",
        prefixesInUse: new Set(["prd", "prdc", "prdct", "prd2"]),
      }),
    ).toBe("prd3");
  });

  it("treats a retired prefix as taken, so an old row ID never points at a new Table", () => {
    expect(
      idPrefixes.fromTitle({
        title: "Item",
        prefixesInUse: new Set(),
        retiredPrefixes: ["itm"],
      }),
    ).toBe("itm2");
    expect(
      idPrefixes.fromTitle({
        title: "Product",
        prefixesInUse: new Set(["prd"]),
        retiredPrefixes: ["prdc"],
      }),
    ).toBe("prdct");
  });

  it("uses s plus a number suffix when the title has no letters", () => {
    expect(
      idPrefixes.fromTitle({ title: "2024", prefixesInUse: new Set() }),
    ).toBe("s");
    expect(
      idPrefixes.fromTitle({ title: "2024", prefixesInUse: new Set(["s"]) }),
    ).toBe("s2");
  });

  it("splits words on punctuation and drops digits before abbreviating", () => {
    expect(
      idPrefixes.fromTitle({ title: "Item 2", prefixesInUse: new Set() }),
    ).toBe("itm");
    expect(
      idPrefixes.fromTitle({ title: "Item-2B!", prefixesInUse: new Set() }),
    ).toBe("ib");
  });

  it("returns only lowercase letters and digits", () => {
    const prefixes = [
      idPrefixes.fromTitle({ title: "Product", prefixesInUse: new Set() }),
      idPrefixes.fromTitle({ title: "Run Item", prefixesInUse: new Set() }),
      idPrefixes.fromTitle({ title: "2024", prefixesInUse: new Set(["s"]) }),
      idPrefixes.fromTitle({ title: "Item-2B!", prefixesInUse: new Set() }),
      idPrefixes.fromTitle({
        title: "Product",
        prefixesInUse: new Set(["prd", "prdc", "prdct"]),
      }),
    ];
    prefixes.forEach((prefix) => {
      expect(prefix).toMatch(/^[a-z0-9]+$/);
    });
  });
});

describe("idPrefixes.retiredAfter", () => {
  it("carries the retired list forward and adds the prefix of each Table gone since the last run", () => {
    expect(
      idPrefixes.retiredAfter({
        previouslyRetired: ["old"],
        previousByTableId: new Map([
          ["tbl-a", "aaa"],
          ["tbl-b", "bbb"],
        ]),
        currentByTableId: new Map([["tbl-a", "aaa"]]),
      }),
    ).toEqual(["bbb", "old"]);
  });

  it("leaves out a prefix a current Table holds, so a Table given back API access keeps its own", () => {
    expect(
      idPrefixes.retiredAfter({
        previouslyRetired: ["aaa"],
        previousByTableId: new Map(),
        currentByTableId: new Map([["tbl-a", "aaa"]]),
      }),
    ).toEqual([]);
  });
});
