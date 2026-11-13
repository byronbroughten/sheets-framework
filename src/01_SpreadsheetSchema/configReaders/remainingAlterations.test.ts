import { describe, expect, it } from "vitest";

import { remainingAlterations } from "./remainingAlterations";
import type { WideAlterations } from "./SyncPlanner/wideAlterations";

function landedTrait(table: string, column: string, path: string) {
  return { table, column, path };
}

describe("remainingAlterations", () => {
  it("removes a landed column trait and keeps the rest", () => {
    const alterations: WideAlterations = {
      tables: {
        item: {
          columns: {
            note: { emptyValueAllowed: false, header: "Note" },
            count: { emptyValueAllowed: true },
          },
        },
        unit: { columns: { size: { emptyValueAllowed: false } } },
      },
    };
    expect(
      remainingAlterations(alterations, [
        landedTrait("item", "note", "emptyValueAllowed"),
        landedTrait("unit", "size", "emptyValueAllowed"),
      ]),
    ).toEqual({
      tables: {
        item: {
          columns: {
            note: { header: "Note" },
            count: { emptyValueAllowed: true },
          },
        },
      },
    });
  });

  it("drops emptied columns, Tables and the tables key", () => {
    const alterations: WideAlterations = {
      tables: {
        item: { columns: { note: { emptyValueAllowed: false } } },
      },
    };
    expect(
      remainingAlterations(alterations, [
        landedTrait("item", "note", "emptyValueAllowed"),
      ]),
    ).toEqual({});
  });

  it("keeps a Table's other traits when its columns empty", () => {
    const alterations: WideAlterations = {
      tables: {
        item: {
          name: "thing",
          columns: { note: { emptyValueAllowed: false } },
        },
      },
    };
    expect(
      remainingAlterations(alterations, [
        landedTrait("item", "note", "emptyValueAllowed"),
      ]),
    ).toEqual({ tables: { item: { name: "thing" } } });
  });

  it("keeps the original key order and other keys", () => {
    const alterations: WideAlterations = {
      tabs: { itemTab: { hidden: true } },
      tables: {
        b: { columns: { y: { emptyValueAllowed: true }, x: { header: "X" } } },
        a: { name: "aa", columns: { z: { emptyValueAllowed: true } } },
      },
    };
    const remaining = remainingAlterations(alterations, [
      landedTrait("b", "y", "emptyValueAllowed"),
    ]);
    expect(Object.keys(remaining)).toEqual(["tabs", "tables"]);
    expect(Object.keys(remaining.tables ?? {})).toEqual(["b", "a"]);
    expect(Object.keys(remaining.tables?.a ?? {})).toEqual(["name", "columns"]);
  });

  it("doesn't mutate its input", () => {
    const alterations: WideAlterations = {
      tables: { item: { columns: { note: { emptyValueAllowed: false } } } },
    };
    remainingAlterations(alterations, [
      landedTrait("item", "note", "emptyValueAllowed"),
    ]);
    expect(alterations).toEqual({
      tables: { item: { columns: { note: { emptyValueAllowed: false } } } },
    });
  });

  it("removes a landed letApiAccess, and the entry once it is empty", () => {
    const alterations = remainingAlterations(
      {
        tables: {
          item: { letApiAccess: true },
          log: { letApiAccess: false, name: "journal" },
        },
      },
      [
        { table: "item", path: "letApiAccess" },
        { table: "log", path: "letApiAccess" },
      ],
    );
    expect(alterations).toEqual({ tables: { log: { name: "journal" } } });
  });

  it("removes a landed Table trait, keeping the Table's columns", () => {
    const remaining = remainingAlterations(
      {
        tables: {
          item: { name: "Stock", columns: { note: { header: "Remark" } } },
        },
      },
      [{ table: "item", path: "name" }],
    );
    expect(remaining).toEqual({
      tables: { item: { columns: { note: { header: "Remark" } } } },
    });
  });

  it("throws on a landed subject that isn't a Table, column or tab trait", () => {
    expect(() => remainingAlterations({}, [{ table: "item" }])).toThrow(
      "Table, column or tab trait",
    );
    expect(() => remainingAlterations({}, [{ tab: "itemTab" }])).toThrow(
      "Table, column or tab trait",
    );
  });
});
