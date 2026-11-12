import { describe, expect, it } from "vitest";

import { assertIdPrefixesUnchanged } from "./idPrefixGuard";

const previousByTableId = new Map([
  ["item", "itm"],
  ["log", "log"],
]);

describe("assertIdPrefixesUnchanged", () => {
  it("passes when every Table keeps its previous prefix", () => {
    expect(() =>
      assertIdPrefixesUnchanged({
        previousByTableId,
        nextByTableId: new Map(previousByTableId),
        retiredPrefixes: new Set(["old"]),
      }),
    ).not.toThrow();
  });

  it("throws naming every Table whose prefix differs from the configs", () => {
    expect(() =>
      assertIdPrefixesUnchanged({
        previousByTableId,
        nextByTableId: new Map([
          ["item", "zzz"],
          ["log", "lgg"],
        ]),
        retiredPrefixes: new Set(),
      }),
    ).toThrow(
      'Table "item" had ID prefix "itm" and is now "zzz". Table "log" had ID prefix "log" and is now "lgg".',
    );
  });

  it("throws for a Table new to the configs whose prefix was retired", () => {
    expect(() =>
      assertIdPrefixesUnchanged({
        previousByTableId,
        nextByTableId: new Map([...previousByTableId, ["gizmo", "old"]]),
        retiredPrefixes: new Set(["old"]),
      }),
    ).toThrow(
      'Table "gizmo" has ID prefix "old", which an earlier Table retired.',
    );
  });

  it("passes a Table new to the configs with an unretired prefix, and ignores a gone one", () => {
    expect(() =>
      assertIdPrefixesUnchanged({
        previousByTableId,
        nextByTableId: new Map([
          ["item", "itm"],
          ["gizmo", "gzm"],
        ]),
        retiredPrefixes: new Set(["log"]),
      }),
    ).not.toThrow();
  });
});
