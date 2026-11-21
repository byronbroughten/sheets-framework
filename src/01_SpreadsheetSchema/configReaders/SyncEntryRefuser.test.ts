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

describe("SyncEntryRefuser, a new tab's before", () => {
  function newTabOf(title: string, before: string) {
    return {
      title,
      before,
      tables: [{ name: `${title} Log`, columns: [{ header: "Note" }] }],
    };
  }

  it("lets through a before naming a mapped tab or another new tab's title", () => {
    expect(
      refusals({
        alterations: {
          newTabs: [
            newTabOf("Ledger", "items"),
            newTabOf("Payments", "Ledger"),
          ],
        },
      }),
    ).toEqual([]);
  });

  it("refuses a before naming neither, by its path and the name", () => {
    expect(
      refusals({
        alterations: {
          newTabs: [
            newTabOf("Ledger", "items"),
            newTabOf("Payments", "ledger"),
          ],
        },
      }),
    ).toEqual([
      {
        kind: "noSuchBeforeTab",
        table: "Payments",
        path: "newTabs[1].before",
        before: "ledger",
      },
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

  it("lets a column's type and dropdown rule through", () => {
    expect(
      refusals({
        alterations: {
          tables: {
            item: {
              columns: {
                name: {
                  columnType: {
                    type: "DROPDOWN",
                    rule: { condition: "ONE_OF_LIST", values: ["a", "b"] },
                  },
                },
                count: { columnType: { type: "CURRENCY", rule: null } },
              },
            },
          },
        },
      }),
    ).toEqual([]);
  });

  it("lets a moves group through", () => {
    expect(
      refusals({
        alterations: {
          tables: {
            item: { moves: [{ columns: ["count"], before: "name" }] },
          },
        },
      }),
    ).toEqual([]);
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

  it("lets a Table's name, deleted, new columns and conditional formats through", () => {
    expect(
      refusals({
        alterations: {
          tables: {
            item: {
              name: "thing",
              deleted: true,
              newColumns: [{ columns: [{ header: "New" }] }],
              conditionalFormats: { add: [], remove: [] },
            },
          },
        },
      }),
    ).toEqual([]);
  });

  it("lets a new column's header, position and empty value allowed through, refusing its other traits", () => {
    expect(
      refusals({
        alterations: {
          tables: {
            item: {
              newColumns: [
                {
                  columns: [{ header: "New", emptyValueAllowed: true }],
                  before: "count",
                },
                {
                  columns: [
                    {
                      header: "Total",
                      formula: "=1",
                      columnType: { type: "TEXT", rule: null },
                    },
                  ],
                },
              ],
            },
          },
        },
      }),
    ).toEqual([
      {
        kind: "notSendableYet",
        table: "item",
        path: "newColumns[1].columns[0].formula",
      },
      {
        kind: "notSendableYet",
        table: "item",
        path: "newColumns[1].columns[0].columnType",
      },
    ]);
  });

  it("lets a tab's traits and new Tables through", () => {
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
              newTables: [
                {
                  name: "Extra",
                  columns: [{ header: "Note", emptyValueAllowed: true }],
                  gap: 1,
                },
              ],
            },
          },
        },
      }),
    ).toEqual([]);
  });

  it("lets a new Table's column formula, type and rule adds through", () => {
    expect(
      refusals({
        alterations: {
          tabs: {
            items: {
              newTables: [
                {
                  name: "Extra",
                  columns: [
                    { header: "Note" },
                    {
                      header: "Count",
                      formula: "=1",
                      columnType: { type: "NUMBER", rule: null },
                    },
                  ],
                  conditionalFormats: {
                    add: [
                      {
                        columns: "all",
                        condition: { type: "NOT_BLANK" },
                        format: {},
                      },
                    ],
                  },
                },
              ],
            },
          },
        },
      }),
    ).toEqual([]);
  });

  it("refuses a new Table's rule text style by the tab's path", () => {
    expect(
      refusals({
        alterations: {
          tabs: {
            items: {
              newTables: [
                {
                  name: "Extra",
                  columns: [{ header: "Note" }],
                  conditionalFormats: {
                    add: [
                      {
                        columns: "all",
                        condition: { type: "NOT_BLANK" },
                        format: { bold: true },
                      },
                    ],
                  },
                },
              ],
            },
          },
        },
      }),
    ).toEqual([
      {
        kind: "notSendableYet",
        table: "items",
        path: "newTables[0].conditionalFormats.add[0].format.bold",
      },
    ]);
  });

  it("lets a new tab's title, traits and Tables through", () => {
    expect(
      refusals({
        alterations: {
          newTabs: [
            {
              title: "Vendors",
              color: "#ff0000",
              before: "items",
              frozenRows: 1,
              frozenColumns: 1,
              hidden: true,
              tables: [
                {
                  name: "Vendor",
                  columns: [
                    { header: "Name" },
                    { header: "Rate", formula: "=1" },
                  ],
                },
              ],
            },
          ],
        },
      }),
    ).toEqual([]);
  });

  it("lets several new Tables on one tab through, and refuses a new tab with none", () => {
    const table = { name: "Extra", columns: [{ header: "Note" }] };
    expect(
      refusals({
        alterations: {
          tabs: { items: { newTables: [table, table] } },
          newTabs: [
            { title: "Vendors", tables: [table, table] },
            { title: "Empty", tables: [] },
          ],
        },
      }),
    ).toEqual([
      { kind: "notSendableYet", table: "Empty", path: "newTabs[1].tables" },
    ]);
  });

  it("refuses a later new Table's unsendable traits at its own index", () => {
    const table = { name: "Extra", columns: [{ header: "Note" }] };
    const bold = {
      ...table,
      conditionalFormats: {
        add: [
          {
            columns: "all" as const,
            condition: { type: "NOT_BLANK" as const },
            format: { bold: true },
          },
        ],
      },
    };
    expect(
      refusals({
        alterations: { tabs: { items: { newTables: [table, bold] } } },
      }),
    ).toEqual([
      {
        kind: "notSendableYet",
        table: "items",
        path: "newTables[1].conditionalFormats.add[0].format.bold",
      },
    ]);
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

describe("SyncEntryRefuser, a rule's columns and braces", () => {
  function ruleRefusals(columns: string[], formula: string) {
    return refusals({
      alterations: {
        tables: {
          item: {
            newColumns: [
              { columns: [{ header: "Late fee" }] },
              { columns: [{ header: "Ref" }] },
              { columns: [{ header: "ref" }] },
            ],
            conditionalFormats: {
              add: [
                {
                  columns,
                  condition: { type: "CUSTOM_FORMULA", values: [formula] },
                  format: { backgroundColor: "#000000" },
                },
              ],
            },
          },
        },
      },
    });
  }

  // The fixture's two ref headers are refused themselves, beside whatever a rule's names add.
  const repeated = {
    kind: "repeatedNewColumn",
    table: "item",
    path: "newColumns[2].columns[0].header",
    firstPath: "newColumns[1].columns[0].header",
    column: "ref",
  };

  it("lets through a rule naming mapped columns and ones the run's newColumns create", () => {
    expect(ruleRefusals(["name", "lateFee"], "=${@lateFee}>${@count}")).toEqual(
      [repeated],
    );
  });

  it("refuses a column or brace naming neither, listing both kinds of column", () => {
    const columns = ["name", "count", "lateFee", "ref"];
    expect(ruleRefusals(["lateFees"], "=${@cost}>0")).toEqual([
      repeated,
      {
        kind: "unknownColumn",
        table: "item",
        path: "conditionalFormats.add[0].columns",
        column: "lateFees",
        columns,
      },
      {
        kind: "unknownBrace",
        table: "item",
        path: "conditionalFormats.add[0].condition.values[0]",
        brace: "${@cost}",
        braceTable: "item",
        columns,
      },
    ]);
  });

  it("leaves a column or brace naming a repeated new header to the repeat's refusal", () => {
    expect(ruleRefusals(["ref"], "=${@ref}>0")).toEqual([repeated]);
  });

  it("checks a replace's rule by its handle's path", () => {
    expect(
      refusals({
        alterations: {
          tables: {
            item: {
              conditionalFormats: {
                replace: {
                  "name | =TRUE": {
                    columns: ["cost"],
                    condition: { type: "NUMBER_GREATER", values: ["${cost}"] },
                    format: {},
                  },
                },
              },
            },
          },
        },
      }),
    ).toEqual([
      {
        kind: "unknownColumn",
        table: "item",
        path: 'conditionalFormats.replace["name | =TRUE"].columns',
        column: "cost",
        columns: ["name", "count"],
      },
    ]);
  });
});

describe("SyncEntryRefuser, entries that can't all hold", () => {
  const rule = {
    columns: ["name"],
    condition: { type: "NOT_BLANK" },
    format: { backgroundColor: "#000000" },
  };

  it("refuses a new column whose header gives the column name an earlier one's does, naming both", () => {
    expect(
      refusals({
        alterations: {
          tables: {
            item: {
              newColumns: [
                { columns: [{ header: "Ref" }] },
                { columns: [{ header: "Note" }] },
                { columns: [{ header: "ref" }] },
              ],
            },
          },
        },
      }),
    ).toEqual([
      {
        kind: "repeatedNewColumn",
        table: "item",
        path: "newColumns[2].columns[0].header",
        firstPath: "newColumns[0].columns[0].header",
        column: "ref",
      },
    ]);
  });

  it("refuses a rule handle both removed and replaced, naming both", () => {
    expect(
      refusals({
        alterations: {
          tables: {
            item: {
              conditionalFormats: {
                remove: ["other", "name | =TRUE"],
                replace: { "name | =TRUE": rule },
              },
            },
          },
        },
      }),
    ).toEqual([
      {
        kind: "removedAndReplaced",
        table: "item",
        path: 'conditionalFormats.replace["name | =TRUE"]',
        removedAt: 'conditionalFormats.remove["name | =TRUE"]',
      },
    ]);
  });

  function renamedRuleRefusals(columns: string[], formula: string) {
    return refusals({
      alterations: {
        tables: {
          item: {
            columns: { count: { header: "Total" } },
            conditionalFormats: {
              add: [
                {
                  columns,
                  condition: { type: "CUSTOM_FORMULA", values: [formula] },
                  format: {},
                },
              ],
            },
          },
        },
      },
    });
  }

  it("refuses a rule naming a column by the name a rename in the same run gives it", () => {
    const renamed = {
      kind: "renamedColumn",
      table: "item",
      column: "total",
      oldName: "count",
    };
    expect(renamedRuleRefusals(["total"], "=${@total}>0")).toEqual([
      { ...renamed, path: "conditionalFormats.add[0].columns" },
      { ...renamed, path: "conditionalFormats.add[0].condition.values[0]" },
    ]);
  });

  it("refuses a brace naming another Table's column by the name its rename in the same run gives it", () => {
    expect(
      refusals({
        alterations: {
          tables: { item: { columns: { count: { header: "Total" } } } },
          newTabs: [
            {
              title: "Orders",
              tables: [
                {
                  name: "order",
                  columns: [{ header: "Ref" }],
                  conditionalFormats: {
                    add: [
                      {
                        columns: ["ref"],
                        condition: {
                          type: "CUSTOM_FORMULA",
                          values: ["=${item[@total]}>0"],
                        },
                        format: {},
                      },
                    ],
                  },
                },
              ],
            },
          ],
        },
      }),
    ).toEqual([
      {
        kind: "renamedColumn",
        table: "Orders",
        path: "newTabs[0].tables[0].conditionalFormats.add[0].condition.values[0]",
        column: "total",
        oldName: "count",
      },
    ]);
  });

  it("lets through a rule naming a renamed column by its old name", () => {
    expect(renamedRuleRefusals(["count"], "=${@count}>0")).toEqual([]);
  });

  it("reads a name two renames swap as the old name it still is", () => {
    expect(
      refusals({
        alterations: {
          tables: {
            item: {
              columns: {
                name: { header: "Count" },
                count: { header: "Name" },
              },
              conditionalFormats: { add: [rule] },
            },
          },
        },
      }),
    ).toEqual([]);
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
