import { describe, expect, it } from "vitest";

import { makeSheetConfigs, makeSpreadsheetConfig } from "./makeConfigs";

const validLayout = {
  idDelimiter: ":",
  idHeader: "ID",
  startTableColIndexBase0: 0,
  columnIdRowIdxBase0: 0,
  columnGroupHeadingRowIndexBase0: 1,
  actionRowIndexBase0: 2,
  tableHeaderRowIndexBase0: 3,
} as const;

describe("makeSpreadsheetConfig", () => {
  it("loads a layout whose uniform-row indexes are distinct and off the first data row", () => {
    expect(makeSpreadsheetConfig(validLayout)).toEqual(validLayout);
  });
});

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
