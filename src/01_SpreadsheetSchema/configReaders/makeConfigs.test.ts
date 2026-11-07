import { describe, expect, it } from "vitest";

import {
  type ColumnConfigStored,
  makeColumnConfigs,
  makeTableConfigs,
} from "./makeConfigs";

describe("makeTableConfigs", () => {
  const tableConfig = {
    tableId: "t1",
    tableName: "Item",
    sheetGid: 1,
    idPrefix: "itm",
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

describe("makeColumnConfigs", () => {
  it("loads an older configs file that records no baseline, leaving each baseline trait unrecorded", () => {
    const columnConfigs = makeColumnConfigs({
      item: {
        amount: {
          columnId: "c1",
          header: "Amount",
          valueName: "number",
          isFormula: false,
          emptyValueAllowed: false,
          customDefaultValue: null,
        },
      },
    });

    const amount: ColumnConfigStored = columnConfigs.item.amount;
    expect([
      amount.formula,
      amount.columnType,
      amount.actionRowCell,
      amount.groupHeading1,
    ]).toEqual([undefined, undefined, undefined, undefined]);
  });
});
