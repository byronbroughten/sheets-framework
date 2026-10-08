import { describe, expect, it } from "vitest";

import { makeTableConfigs } from "./makeConfigs";

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
