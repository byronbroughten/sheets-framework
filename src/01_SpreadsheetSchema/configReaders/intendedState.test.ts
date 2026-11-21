import { describe, expect, it } from "vitest";

import { assertType, type IsExactly } from "../../testSupport/typeAssertions";
import {
  makePendingAlterations,
  type PendingAlterations,
} from "./intendedState";

type TableEntries = NonNullable<PendingAlterations["tables"]>;
type ItemEntry = NonNullable<TableEntries["item"]>;
type ItemColumnEntries = NonNullable<ItemEntry["columns"]>;
type ItemNameEntry = NonNullable<ItemColumnEntries["name"]>;
type ItemNewColumn = NonNullable<ItemEntry["newColumns"]>[number];
type ItemMove = NonNullable<ItemEntry["moves"]>[number];
type ItemRuleColumn = Exclude<
  NonNullable<
    NonNullable<ItemEntry["conditionalFormats"]>["add"]
  >[number]["columns"],
  "all"
>[number];
type ItemRule = NonNullable<
  NonNullable<ItemEntry["conditionalFormats"]>["add"]
>[number];
type ItemColumnName = "optionalNote" | "requiredCount" | "id" | "name";
type TabEntries = NonNullable<PendingAlterations["tabs"]>;
type ItemTabEntry = NonNullable<TabEntries["item"]>;
type NewTabs = NonNullable<PendingAlterations["newTabs"]>;

describe("makePendingAlterations", () => {
  it("hands back the alterations it was given", () => {
    const alterations = { tables: { item: { letApiAccess: true } } };
    expect(makePendingAlterations(alterations)).toBe(alterations);
  });

  it("takes every Table, column and tab trait, keyed by codebase name", () => {
    makePendingAlterations({
      tables: {
        item: {
          name: "stockItem",
          letApiAccess: true,
          columns: {
            optionalNote: {
              header: "Note",
              columnType: {
                type: "DROPDOWN",
                rule: { condition: "ONE_OF_LIST", values: ["a", "b"] },
              },
              emptyValueAllowed: true,
            },
            requiredCount: { formula: "=${log[amount]} * 2" },
            id: { deleted: true },
          },
          newColumns: [
            {
              columns: [{ header: "Price", emptyValueAllowed: true }],
              before: "requiredCount",
            },
          ],
          moves: [{ columns: ["optionalNote"], before: "name" }],
        },
        log: {
          conditionalFormats: {
            add: [
              {
                columns: ["amount"],
                condition: { type: "CUSTOM_FORMULA", values: ["=${amount}<0"] },
                format: { foregroundColor: "#cc0000", bold: true },
              },
            ],
            remove: ["log-amount-1"],
            replace: {
              "log-entry-1": {
                columns: "all",
                condition: { type: "TEXT_EQ", values: ["done"] },
                format: { backgroundColor: "#eeeeee" },
              },
            },
          },
        },
        layoutRight: { deleted: true },
      },
      tabs: {
        valueTypes: {
          title: "Values",
          color: "#3366ff",
          before: "item",
          frozenRows: 4,
          frozenColumns: 1,
          hidden: false,
          newTables: [{ name: "extra", columns: [{ header: "Note" }], gap: 1 }],
        },
        log: { deleted: true },
      },
      newTabs: [
        {
          title: "Fresh",
          before: "layout",
          tables: [{ name: "fresh", columns: [{ header: "Name" }] }],
        },
      ],
    });
  });

  it("is empty with no alterations at all", () => {
    expect(makePendingAlterations({})).toEqual({});
  });
});

describe("PendingAlterations' keys", () => {
  it("are exactly the map's names", () => {
    assertType<IsExactly<keyof ItemColumnEntries, ItemColumnName>>(true);
    assertType<IsExactly<ItemMove["before"], ItemColumnName>>(true);
    assertType<IsExactly<ItemMove["columns"][number], ItemColumnName>>(true);
  });

  it("reject a name the ID-name map doesn't hold", () => {
    // @ts-expect-error no Table is named notATable
    const unknownTable: TableEntries = { notATable: {} };
    // @ts-expect-error item has no column notAColumn
    const unknownColumn: ItemColumnEntries = { notAColumn: {} };
    // @ts-expect-error no tab is named notATab
    const unknownTab: TabEntries = { notATab: {} };
    expect([unknownTable, unknownColumn, unknownTab]).toHaveLength(3);
  });

  it("name a Table by its codebase name, never its ID", () => {
    assertType<"scratch208an" extends keyof TableEntries ? true : false>(true);
    // @ts-expect-error tbl-208an is the scratch208an Table's ID
    const byId: TableEntries = { "tbl-208an": {} };
    expect(byId).toBeDefined();
  });

  it("name a tab by its tab name, never its title", () => {
    assertType<"valueTypes" extends keyof TabEntries ? true : false>(true);
    // @ts-expect-error "Value Types" is the tab's title, display text
    const byTitle: TabEntries = { "Value Types": {} };
    expect(byTitle).toBeDefined();
  });

  it("place a column only before another column of its own Table", () => {
    // @ts-expect-error entry is a log column, not an item one
    const otherTable: ItemMove = { columns: ["id"], before: "entry" };
    const newColumn: ItemNewColumn = {
      columns: [{ header: "X" }],
      // @ts-expect-error a new column group too
      before: "entry",
    };
    expect([otherTable, newColumn]).toHaveLength(2);
  });

  it("place an existing column only through moves", () => {
    // @ts-expect-error a column entry takes no before
    const ownBefore: ItemNameEntry = { before: "id" };
    expect(ownBefore).toBeDefined();
  });

  it("let a Table's rules name a column its newColumns create, by the name its header gives", () => {
    assertType<
      IsExactly<
        ItemRuleColumn,
        ItemColumnName | (string & Record<never, never>)
      >
    >(true);
    const overNewColumn: ItemEntry = {
      newColumns: [{ columns: [{ header: "Late fee" }] }],
      conditionalFormats: {
        add: [
          {
            columns: ["lateFee"],
            condition: { type: "CUSTOM_FORMULA", values: ["=${@lateFee}>0"] },
            format: {},
          },
        ],
      },
    };
    expect(overNewColumn).toBeDefined();
  });

  it("place a tab only before a tab the map holds", () => {
    // @ts-expect-error no tab is named notATab
    const unknownBefore: ItemTabEntry = { before: "notATab" };
    expect(unknownBefore).toBeDefined();
  });

  it("give a new column and a new tab their display text", () => {
    const headerless: ItemEntry = {
      // @ts-expect-error a new column needs its header
      newColumns: [{ columns: [{ emptyValueAllowed: true }] }],
    };
    // @ts-expect-error a new tab needs its title
    const untitled: NewTabs = [{ tables: [] }];
    expect([headerless, untitled]).toHaveLength(2);
  });
});

describe("a rule's condition", () => {
  function ruleOf(condition: ItemRule["condition"]): ItemRule {
    return { columns: "all", condition, format: {} };
  }

  it("takes a number as the one value of a NUMBER_* type", () => {
    const rules = [
      ruleOf({ type: "NUMBER_EQ", values: ["0"] }),
      ruleOf({ type: "NUMBER_GREATER", values: ["1.50"] }),
      ruleOf({ type: "NUMBER_LESS", values: ["-2"] }),
    ];
    expect(rules).toHaveLength(3);
  });

  it("refuses a boolean or a word as the value of a NUMBER_* type", () => {
    // @ts-expect-error a boolean's text isn't a number
    const boolean = ruleOf({ type: "NUMBER_EQ", values: ["TRUE"] });
    // @ts-expect-error nor is Yes
    const word = ruleOf({ type: "NUMBER_NOT_EQ", values: ["Yes"] });
    // @ts-expect-error a NUMBER_* type takes one value
    const two = ruleOf({ type: "NUMBER_LESS", values: ["1", "2"] });
    expect([boolean, word, two]).toHaveLength(3);
  });

  it("leaves TEXT_EQ and CUSTOM_FORMULA their string values", () => {
    const rules = [
      ruleOf({ type: "TEXT_EQ", values: ["TRUE"] }),
      ruleOf({ type: "CUSTOM_FORMULA", values: ["=${@count}>0"] }),
    ];
    expect(rules).toHaveLength(2);
  });

  it("types a NUMBER_* value as a number string and the others as a string", () => {
    type Condition = ItemRule["condition"];
    type NumberValues = (Condition & { type: "NUMBER_EQ" })["values"];
    type TextValues = (Condition & { type: "TEXT_EQ" })["values"];
    assertType<IsExactly<NumberValues, [`${number}`] | undefined>>(true);
    assertType<IsExactly<TextValues, string[] | undefined>>(true);
  });
});

describe("deleted: true", () => {
  it("stands alone in a Table, column or tab entry", () => {
    assertType<{ deleted: true } extends ItemEntry ? true : false>(true);
    assertType<{ deleted: true } extends ItemNameEntry ? true : false>(true);
    assertType<{ deleted: true } extends ItemTabEntry ? true : false>(true);
  });

  it("beside any other trait is a type error", () => {
    // @ts-expect-error a deleted Table takes no rename
    const renamedTable: ItemEntry = { deleted: true, name: "x" };
    // @ts-expect-error nor a column entry
    const tableColumns: ItemEntry = { deleted: true, columns: {} };
    // @ts-expect-error a deleted column takes no header
    const column: ItemNameEntry = { deleted: true, header: "Name" };
    // @ts-expect-error a deleted tab takes no title
    const tab: ItemTabEntry = { deleted: true, title: "Items" };
    // @ts-expect-error nor a new Table
    const tabTables: ItemTabEntry = { deleted: true, newTables: [] };
    expect([renamedTable, tableColumns, column, tab, tabTables]).toHaveLength(
      5,
    );
  });
});
