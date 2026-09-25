import { describe, expect, it } from "vitest";

import { makeSheetConfigs } from "./makeConfigs";

describe("makeSheetConfigs", () => {
  it("throws when two sheets share a non-empty ID prefix", () => {
    expect(() =>
      makeSheetConfigs({
        item: {
          sheetGid: 1,
          idPrefix: "itm",
          hasIdColumn: true,
          hasNameColumn: false,
        },
        runItem: {
          sheetGid: 2,
          idPrefix: "itm",
          hasIdColumn: true,
          hasNameColumn: false,
        },
      }),
    ).toThrow(/item.*runItem.*"itm"/);
  });

  it("throws when a sheet has an empty ID prefix", () => {
    expect(() =>
      makeSheetConfigs({
        notes: {
          sheetGid: 1,
          idPrefix: "",
          hasIdColumn: false,
          hasNameColumn: false,
        },
      }),
    ).toThrow(/notes.*no ID prefix/);
  });
});
