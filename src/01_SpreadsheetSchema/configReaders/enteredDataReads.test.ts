import { describe, expect, it } from "vitest";

import { enteredDataReads } from "./enteredDataReads";
import type { IdNameMapGeneric } from "./idNameMapTypes";

const idNameMap: IdNameMapGeneric = {
  tables: {
    "t-item": {
      name: "item",
      sheetGid: 11,
      columns: { note: "c-note", total: "c-total" },
    },
    "t-log": { name: "log", sheetGid: 22 },
  },
  tabs: {
    "11": { name: "itemTab", title: "Item", tableIds: ["t-item"] },
    "22": { name: "logTab", title: "Log", tableIds: ["t-log"] },
  },
};

describe("enteredDataReads", () => {
  it("reads nothing when nothing is deleted and no formula is sent", () => {
    expect(
      enteredDataReads({
        alterations: {
          tables: { item: { columns: { note: { header: "Memo" } } } },
        },
        sentTraits: [
          {
            kind: "column",
            table: "item",
            column: "note",
            trait: "header",
            value: "Memo",
          },
        ],
        idNameMap,
      }),
    ).toEqual({ tabGids: [], tableIds: [], columns: [] });
  });

  it("reads each deleted tab, Table and column by its ID", () => {
    expect(
      enteredDataReads({
        alterations: {
          tables: {
            item: { columns: { note: { deleted: true } } },
            log: { deleted: true },
          },
          tabs: { logTab: { deleted: true } },
        },
        sentTraits: [],
        idNameMap,
      }),
    ).toEqual({
      tabGids: [22],
      tableIds: ["t-log"],
      columns: [{ tableId: "t-item", columnId: "c-note" }],
    });
  });

  it("reads a column whose sent trait is a formula", () => {
    expect(
      enteredDataReads({
        alterations: {},
        sentTraits: [
          {
            kind: "column",
            table: "item",
            column: "total",
            trait: "formula",
            value: "={note}",
          },
        ],
        idNameMap,
      }),
    ).toEqual({
      tabGids: [],
      tableIds: [],
      columns: [{ tableId: "t-item", columnId: "c-total" }],
    });
  });

  it("skips a name the map can't resolve to an ID", () => {
    expect(
      enteredDataReads({
        alterations: {
          tables: { log: { columns: { body: { deleted: true } } } },
          tabs: { goneTab: { deleted: true } },
        },
        sentTraits: [],
        idNameMap,
      }),
    ).toEqual({ tabGids: [], tableIds: [], columns: [] });
  });
});
