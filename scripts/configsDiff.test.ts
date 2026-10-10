import { describe, expect, it } from "vitest";

import { type ConfigSnapshot, summarizeConfigsDiff } from "./configsDiff.ts";

const item = { tableId: "item", tableName: "item", hasIdColumn: true };
const itemName = { columnId: "c:itm:aaa", header: "Name", valueName: "string" };
const itemQty = { columnId: "c:itm:bbb", header: "Qty", valueName: "number" };

function snapshot(overrides: Partial<ConfigSnapshot> = {}): ConfigSnapshot {
  return {
    tableConfigs: { item },
    columnConfigs: { item: { name: itemName, qty: itemQty } },
    valueConfigs: { status: ["Open", "Closed"] },
    ...overrides,
  };
}

describe("summarizeConfigsDiff", () => {
  it("says no changes when nothing changed", () => {
    expect(
      summarizeConfigsDiff({ before: snapshot(), after: snapshot() }),
    ).toEqual(["No config changes."]);
  });

  it("lists added and removed Tables by name", () => {
    const log = { tableId: "log", tableName: "log", hasIdColumn: false };
    const after = snapshot({
      tableConfigs: { log },
      columnConfigs: { log: {} },
    });
    expect(summarizeConfigsDiff({ before: snapshot(), after })).toEqual([
      "Table item: removed",
      "  column name: removed",
      "  column qty: removed",
      "Table log: added",
      "4 changes.",
    ]);
  });

  it("reports a Table with the same tableId under a new key as one rename", () => {
    const after = snapshot({
      tableConfigs: { goods: item },
      columnConfigs: { goods: { name: itemName, qty: itemQty } },
    });
    expect(summarizeConfigsDiff({ before: snapshot(), after })).toEqual([
      "Table item → goods: renamed",
      "1 change.",
    ]);
  });

  it("lists added, removed and renamed columns under their Table", () => {
    const itemNote = {
      columnId: "c:itm:ccc",
      header: "Note",
      valueName: "string",
    };
    const after = snapshot({
      columnConfigs: { item: { title: itemName, note: itemNote } },
    });
    expect(summarizeConfigsDiff({ before: snapshot(), after })).toEqual([
      "Table item:",
      "  column name → title: renamed",
      "  column qty: removed",
      "  column note: added",
      "3 changes.",
    ]);
  });

  it("reports changed fields on a Table and on a column with old and new values", () => {
    const after = snapshot({
      tableConfigs: { item: { ...item, hasIdColumn: false } },
      columnConfigs: {
        item: { name: itemName, qty: { ...itemQty, valueName: "string" } },
      },
    });
    expect(summarizeConfigsDiff({ before: snapshot(), after })).toEqual([
      "Table item: hasIdColumn true → false",
      '  column qty: valueName "number" → "string"',
      "2 changes.",
    ]);
  });

  it("reports value configs added, removed and changed", () => {
    const before = snapshot({
      valueConfigs: { status: ["Open", "Closed"], size: ["S"] },
    });
    const after = snapshot({
      valueConfigs: { status: ["Open", "Pending"], color: ["Red"] },
    });
    expect(summarizeConfigsDiff({ before, after })).toEqual([
      'Value status: added "Pending"; removed "Closed"',
      "Value size: removed",
      "Value color: added",
      "3 changes.",
    ]);
  });

  it("counts everything as added on a first generation with no HEAD version", () => {
    expect(
      summarizeConfigsDiff({ before: undefined, after: snapshot() }),
    ).toEqual([
      "Table item: added",
      "  column name: added",
      "  column qty: added",
      "Value status: added",
      "4 changes.",
    ]);
  });
});
