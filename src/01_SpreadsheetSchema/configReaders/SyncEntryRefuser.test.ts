import { describe, expect, it } from "vitest";

import type { IdNameMapGeneric } from "./idNameMapTypes";
import {
  SyncEntryRefuser,
  type SyncEntryRefuserProps,
} from "./SyncEntryRefuser";

const idNameMap: IdNameMapGeneric = {
  tables: {
    "t:itm:aaaaaaa": {
      name: "item",
      sheetGid: 1,
      columns: { name: "c:itm:bbbbbbb", count: "c:itm:ccccccc" },
    },
    // A Table with no API access still has its entry in the map.
    "t:log:ddddddd": { name: "log", sheetGid: 2 },
  },
  tabs: {
    "1": { name: "items", title: "Items", tableIds: ["t:itm:aaaaaaa"] },
    "2": { name: "logs", title: "Logs", tableIds: ["t:log:ddddddd"] },
  },
};

function refusals(props: Partial<SyncEntryRefuserProps>) {
  return SyncEntryRefuser.init({
    alterations: {},
    idNameMap,
    ...props,
  }).refusals;
}

describe("SyncEntryRefuser, a Table or tab that no longer exists", () => {
  it("refuses a tables entry the map holds no Table named", () => {
    expect(
      refusals({
        alterations: {
          tables: {
            vendor: { columns: { name: { emptyValueAllowed: true } } },
          },
        },
      }),
    ).toEqual([{ kind: "noSuchTable", entry: "tables", name: "vendor" }]);
  });

  it("refuses a tabs entry the map holds no tab named", () => {
    expect(
      refusals({ alterations: { tabs: { vendors: { hidden: true } } } }),
    ).toEqual([{ kind: "noSuchTable", entry: "tabs", name: "vendors" }]);
  });

  it("names only the missing Table for an entry with other traits too", () => {
    expect(
      refusals({
        alterations: { tables: { vendor: { name: "supplier" } } },
      }),
    ).toEqual([{ kind: "noSuchTable", entry: "tables", name: "vendor" }]);
  });

  it("refuses each missing entry beside the ones that exist", () => {
    expect(
      refusals({
        alterations: {
          tables: {
            vendor: { deleted: true },
            item: { columns: { name: { emptyValueAllowed: false } } },
          },
          tabs: { gone: {} },
        },
      }),
    ).toEqual([
      { kind: "noSuchTable", entry: "tables", name: "vendor" },
      { kind: "noSuchTable", entry: "tabs", name: "gone" },
    ]);
  });
});

describe("SyncEntryRefuser, traits not sendable yet", () => {
  it("lets a Table's name and a column's header and emptyValueAllowed through", () => {
    expect(
      refusals({
        alterations: {
          tables: {
            item: {
              name: "Stock",
              columns: {
                name: { header: "Mine", emptyValueAllowed: true },
                count: { emptyValueAllowed: false },
              },
            },
          },
        },
      }),
    ).toEqual([]);
  });

  it("refuses a column's other traits, naming the Table and the column's path", () => {
    expect(
      refusals({
        alterations: {
          tables: {
            item: {
              columns: {
                name: { columnType: { type: "TEXT", rule: null } },
                count: { formula: "=1", before: "name" },
              },
            },
          },
        },
      }),
    ).toEqual([
      { kind: "notSendableYet", table: "item", path: "name.columnType" },
      { kind: "notSendableYet", table: "item", path: "count.before" },
    ]);
  });

  it("lets a column's deletion through", () => {
    expect(
      refusals({
        alterations: {
          tables: { item: { columns: { count: { deleted: true } } } },
        },
      }),
    ).toEqual([]);
  });

  it("refuses a Table's own traits but its name and deleted", () => {
    expect(
      refusals({
        alterations: {
          tables: {
            item: {
              name: "thing",
              deleted: true,
              newColumns: [{ header: "New" }],
              conditionalFormats: { add: [] },
            },
          },
        },
      }),
    ).toEqual([
      { kind: "notSendableYet", table: "item", path: "newColumns" },
      { kind: "notSendableYet", table: "item", path: "conditionalFormats" },
    ]);
  });

  it("lets a tab's traits through, refusing only its new Tables", () => {
    expect(
      refusals({
        alterations: {
          tabs: {
            items: {
              title: "Things",
              color: "#ff0000",
              before: "items",
              frozenRows: 2,
              frozenColumns: 1,
              hidden: true,
              newTables: [],
            },
          },
        },
      }),
    ).toEqual([{ kind: "notSendableYet", table: "items", path: "newTables" }]);
  });

  it("refuses a new tab, named by its title", () => {
    expect(
      refusals({
        alterations: { newTabs: [{ title: "Vendors", tables: [] }] },
      }),
    ).toEqual([{ kind: "notSendableYet", table: "Vendors", path: "newTabs" }]);
  });

  it("refuses two untitled new tabs once", () => {
    expect(
      refusals({
        alterations: { newTabs: [{ tables: [] }, { tables: [] }] },
      }),
    ).toEqual([{ kind: "notSendableYet", table: "newTabs", path: "newTabs" }]);
  });

  it("refuses nothing for no intended state", () => {
    expect(refusals({})).toEqual([]);
  });
});

describe("SyncEntryRefuser, a formula's braces", () => {
  function formulaRefusals(formula: string) {
    return refusals({
      alterations: { tables: { item: { columns: { count: { formula } } } } },
    });
  }

  it("lets a formula through whose braces all name a column in the map", () => {
    expect(
      formulaRefusals('=${@name}&" "&COUNTA(${name})&"${nothing}"'),
    ).toEqual([]);
  });

  it("refuses a brace naming no column of its Table, listing the Table's columns", () => {
    expect(formulaRefusals("=${@nam}*2")).toEqual([
      {
        kind: "unknownBrace",
        table: "item",
        path: "count.formula",
        brace: "${@nam}",
        braceTable: "item",
        columns: ["name", "count"],
      },
    ]);
  });

  it("refuses a brace naming no Table, with no columns to list", () => {
    expect(formulaRefusals("=SUM(${vendor[@cost]})")).toEqual([
      {
        kind: "unknownBrace",
        table: "item",
        path: "count.formula",
        brace: "${vendor[@cost]}",
        braceTable: "vendor",
        columns: undefined,
      },
    ]);
  });

  it("refuses a misplaced @, which reads as no brace", () => {
    expect(formulaRefusals("=${@item[name]}")).toEqual([
      {
        kind: "malformedBrace",
        table: "item",
        path: "count.formula",
        brace: "${@item[name]}",
      },
    ]);
  });

  it("refuses a brace naming a column of a Table without API access, whose map entry holds none", () => {
    expect(formulaRefusals("=${log[when]}")).toEqual([
      {
        kind: "unknownBrace",
        table: "item",
        path: "count.formula",
        brace: "${log[when]}",
        braceTable: "log",
        columns: [],
      },
    ]);
  });
});

describe("SyncEntryRefuser, API access", () => {
  it("lets a grant and a revoke through", () => {
    expect(
      refusals({
        alterations: {
          tables: {
            log: { letApiAccess: true },
            item: { letApiAccess: false },
          },
        },
      }),
    ).toEqual([]);
  });

  it("refuses every entry on a Table without API access, other than a grant", () => {
    expect(
      refusals({
        alterations: {
          tables: {
            log: {
              name: "journal",
              columns: { x: { emptyValueAllowed: true } },
            },
          },
        },
      }),
    ).toEqual([
      { kind: "noApiAccess", table: "log", path: "name" },
      { kind: "noApiAccess", table: "log", path: "x.emptyValueAllowed" },
    ]);
  });

  it("refuses any other entry on a Table being granted, and allows other Tables' entries beside it", () => {
    expect(
      refusals({
        alterations: {
          tables: {
            log: {
              letApiAccess: true,
              columns: { x: { emptyValueAllowed: true } },
            },
            item: { columns: { name: { emptyValueAllowed: true } } },
          },
        },
      }),
    ).toEqual([
      { kind: "besideGrant", table: "log", path: "x.emptyValueAllowed" },
    ]);
  });
});
