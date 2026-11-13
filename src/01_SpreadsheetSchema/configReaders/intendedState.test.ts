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
              before: "name",
            },
            requiredCount: { formula: "=${log[amount]} * 2" },
            id: { deleted: true },
          },
          newColumns: [
            {
              header: "Price",
              before: "requiredCount",
              emptyValueAllowed: true,
            },
          ],
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
                condition: { type: "NOT_BLANK" },
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
    assertType<IsExactly<NonNullable<ItemNameEntry["before"]>, ItemColumnName>>(
      true,
    );
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
    assertType<"tableConfig" extends keyof TableEntries ? true : false>(true);
    // @ts-expect-error sheetConfig is the Table Config Table's ID
    const byId: TableEntries = { sheetConfig: {} };
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
    const otherTable: ItemNameEntry = { before: "entry" };
    // @ts-expect-error a new column too
    const newColumn: ItemNewColumn = { header: "X", before: "entry" };
    expect([otherTable, newColumn]).toHaveLength(2);
  });

  it("place a tab only before a tab the map holds", () => {
    // @ts-expect-error no tab is named notATab
    const unknownBefore: ItemTabEntry = { before: "notATab" };
    expect(unknownBefore).toBeDefined();
  });

  it("give a new column and a new tab their display text", () => {
    // @ts-expect-error a new column needs its header
    const headerless: ItemEntry = { newColumns: [{ emptyValueAllowed: true }] };
    // @ts-expect-error a new tab needs its title
    const untitled: NewTabs = [{ tables: [] }];
    expect([headerless, untitled]).toHaveLength(2);
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
    // @ts-expect-error nor a place
    const movedColumn: ItemNameEntry = { deleted: true, before: "id" };
    // @ts-expect-error a deleted tab takes no title
    const tab: ItemTabEntry = { deleted: true, title: "Items" };
    // @ts-expect-error nor a new Table
    const tabTables: ItemTabEntry = { deleted: true, newTables: [] };
    expect([
      renamedTable,
      tableColumns,
      column,
      movedColumn,
      tab,
      tabTables,
    ]).toHaveLength(6);
  });
});
