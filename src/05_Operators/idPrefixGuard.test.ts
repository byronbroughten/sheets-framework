import { describe, expect, it } from "vitest";

import { IdPrefixesByTableId } from "../01_SpreadsheetSchema/idPrefixes";
import { validateIdPrefixesUnchanged } from "./idPrefixGuard";

const previousEntries: [string, string][] = [
  ["item", "itm"],
  ["log", "log"],
];
const previousByTableId = new IdPrefixesByTableId(new Map(previousEntries));

describe("validateIdPrefixesUnchanged", () => {
  it("passes when every Table keeps its previous prefix", () => {
    expect(() =>
      validateIdPrefixesUnchanged({
        previousByTableId,
        nextByTableId: new IdPrefixesByTableId(new Map(previousEntries)),
        retiredPrefixes: new Set(["old"]),
      }),
    ).not.toThrow();
  });

  it("throws naming every Table whose prefix differs from the configs", () => {
    expect(() =>
      validateIdPrefixesUnchanged({
        previousByTableId,
        nextByTableId: new IdPrefixesByTableId(
          new Map([
            ["item", "zzz"],
            ["log", "lgg"],
          ]),
        ),
        retiredPrefixes: new Set(),
      }),
    ).toThrow(
      'Table "item" had ID prefix "itm" and is now "zzz". Table "log" had ID prefix "log" and is now "lgg".',
    );
  });

  it("throws for a Table new to the configs whose prefix was retired", () => {
    expect(() =>
      validateIdPrefixesUnchanged({
        previousByTableId,
        nextByTableId: new IdPrefixesByTableId(
          new Map([...previousEntries, ["gizmo", "old"]]),
        ),
        retiredPrefixes: new Set(["old"]),
      }),
    ).toThrow(
      'Table "gizmo" has ID prefix "old", which an earlier Table retired.',
    );
  });

  it("passes a Table new to the configs with an unretired prefix, and ignores a gone one", () => {
    expect(() =>
      validateIdPrefixesUnchanged({
        previousByTableId,
        nextByTableId: new IdPrefixesByTableId(
          new Map([
            ["item", "itm"],
            ["gizmo", "gzm"],
          ]),
        ),
        retiredPrefixes: new Set(["log"]),
      }),
    ).not.toThrow();
  });
});
