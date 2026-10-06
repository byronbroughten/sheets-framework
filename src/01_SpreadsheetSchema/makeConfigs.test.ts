import { describe, expect, it } from "vitest";

import { makeSheetConfigs, makeTableConfigs } from "./makeConfigs";

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

describe("makeTableConfigs", () => {
  const tableConfig = {
    tableId: "t1",
    tableName: "Item",
    sheetGid: 1,
    idPrefix: "itm",
    headerRowIndex: 3,
    startColIndex: 0,
    hasIdColumn: true,
    hasNameColumn: false,
  };

  it("throws when two Tables share a non-empty ID prefix", () => {
    expect(() =>
      makeTableConfigs({
        item: tableConfig,
        runItem: { ...tableConfig, tableId: "t2", tableName: "Run item" },
      }),
    ).toThrow(/item.*runItem.*"itm"/);
  });

  it("throws when a Table has an empty ID prefix", () => {
    expect(() =>
      makeTableConfigs({ notes: { ...tableConfig, idPrefix: "" } }),
    ).toThrow(/notes.*no ID prefix/);
  });
});
