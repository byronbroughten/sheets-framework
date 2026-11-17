import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import type { IdNameMapGeneric } from "../01_SpreadsheetSchema/configReaders/idNameMapTypes";
import { expectedSheetLayout } from "../testSupport/expectedSheetLayout";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import type { FakeSheetsService } from "../testSupport/fakeSheetsService";
import { parseIdNameMapFileSource } from "./parseIdNameMapFileSource";
import type { ConfigRegeneration } from "./SyncCoordinator";
import {
  itemColumns,
  itemGid,
  runItem,
  seedFixture,
} from "./syncCoordinatorTestSupport";

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("SyncCoordinator.run new tabs and Tables", () => {
  const mintedTableId = /^tbl-[0-9a-f]{10}$/;
  const { colIdRowIndex, tableHeaderRowIndex, startTableColIndex } =
    expectedSheetLayout;
  const paymentLog = {
    name: "Payment Log",
    columns: [
      { header: "Payee" },
      { header: "Amount", emptyValueAllowed: true },
    ],
  };

  function mapOf(files: ConfigRegeneration | undefined): IdNameMapGeneric {
    return parseIdNameMapFileSource(
      Val.assert(files, "emitted files").idNameMap,
    );
  }

  function tabGidOf(map: IdNameMapGeneric, title: string): number {
    const found = Object.entries(map.tabs).find(
      ([, tab]) => tab.title === title,
    );
    return Number(Val.assert(found, `mapped tab ${title}`)[0]);
  }

  function itemTableEnd(service: FakeSheetsService): number {
    const range = service.grid.sheet(itemGid).tables[0]?.range;
    return Val.assert(range?.endColumnIndex, "Item Table end");
  }

  it("plans a new tab and each Table on a dry run, sending nothing", () => {
    const service = seedFixture();

    const { result } = runItem({
      isSending: false,
      alterations: {
        newTabs: [{ title: "Payments", tables: [paymentLog] }],
        tabs: {
          item: {
            newTables: [
              { name: "Desk Log", columns: [{ header: "Note" }], gap: 1 },
            ],
          },
        },
      },
    });

    expect(service.batchUpdateCount()).toBe(0);
    expect(result).toMatchObject({
      outcome: { kind: "ready" },
      plan: [
        {
          subject: { tab: "Payments", path: "newTabs" },
          what: "create the tab",
        },
        {
          subject: { tab: "Payments", path: "newTabs.tables" },
          what: 'create Table "Payment Log"',
        },
        {
          subject: { tab: "item", path: "newTables" },
          what: 'create Table "Desk Log" right of the tab\'s Tables, one empty column between',
        },
      ],
    });
  });

  it("adds a new tab holding its Table, with API access on in the configs and the ID-name map, and lists it as created", () => {
    const service = seedFixture();

    const { result, files } = runItem({
      isSending: true,
      alterations: { newTabs: [{ title: "Payments", tables: [paymentLog] }] },
    });

    const emitted = Val.assert(files, "emitted files");
    const map = mapOf(files);
    const gid = tabGidOf(map, "Payments");
    expect(service.grid.sheetTitles()).toEqual(["Item", "Payments"]);
    const [table, ...others] = service.grid.sheet(gid).tables;
    expect(others).toEqual([]);
    const tableId = String(table?.tableId);
    expect(tableId).toMatch(mintedTableId);
    expect(table).toMatchObject({
      name: "Payment Log",
      range: {
        startRowIndex: tableHeaderRowIndex,
        startColumnIndex: startTableColIndex,
        endColumnIndex: startTableColIndex + 2,
      },
    });
    const columnIds = service.grid
      .sheet(gid)
      .values({
        startRowIndex: colIdRowIndex,
        endRowIndex: colIdRowIndex + 1,
        startColumnIndex: startTableColIndex,
        endColumnIndex: startTableColIndex + 2,
      })[0]
      ?.map(String);
    const [payeeId, amountId] = Val.assert(columnIds, "column IDs");
    expect(payeeId).toMatch(/^c:plg:/);
    expect(emitted.tableConfigs).toContain(
      `"paymentLog": { "tableId": "${tableId}"`,
    );
    expect(emitted.columnConfigs).toContain(`"columnId": "${payeeId}"`);
    expect(emitted.columnConfigs).toMatch(
      new RegExp(
        `"amount": \\{[^\\n]*"columnId": "${amountId}"[^\\n]*"emptyValueAllowed": true`,
      ),
    );
    expect(map.tables[tableId]).toEqual({
      name: "paymentLog",
      sheetGid: gid,
      columns: { payee: payeeId, amount: amountId },
    });
    expect(emitted.alterations).toContain("makePendingAlterations({})");
    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      landed: [
        { kind: "send", subject: { tab: "Payments", path: "newTabs" } },
        {
          kind: "createdTable",
          table: "paymentLog",
          tab: "Payments",
          tableId,
          columns: [
            { header: "Payee", column: "payee", columnId: payeeId },
            { header: "Amount", column: "amount", columnId: amountId },
          ],
        },
      ],
    });
  });

  it.each([0, 1] as const)(
    "adds a Table right of Item with gap %i, empty columns between",
    (gap) => {
      const service = seedFixture();
      const itemEnd = itemTableEnd(service);

      const { result, files } = runItem({
        isSending: true,
        alterations: {
          tabs: { item: { newTables: [{ ...paymentLog, gap }] } },
        },
      });

      const tables = service.grid.sheet(itemGid).tables;
      expect(tables.map(({ name }) => name)).toEqual(["Item", "Payment Log"]);
      expect(tables[1]?.range).toMatchObject({
        startRowIndex: tableHeaderRowIndex,
        startColumnIndex: itemEnd + gap,
        endColumnIndex: itemEnd + gap + 2,
      });
      const tableId = String(tables[1]?.tableId);
      expect(mapOf(files).tables[tableId]).toMatchObject({
        name: "paymentLog",
        sheetGid: itemGid,
      });
      expect(mapOf(files).tabs[String(itemGid)]?.tableIds).toEqual([
        "item",
        tableId,
      ]);
      expect(files?.alterations).toContain("makePendingAlterations({})");
      expect(result).toMatchObject({
        landed: [
          { kind: "createdTable", table: "paymentLog", tab: "Item", tableId },
        ],
      });
    },
  );

  it("stops before sending on a new Table whose name a live Table gives, naming it", () => {
    const service = seedFixture();

    const { result, files } = runItem({
      isSending: true,
      alterations: {
        tabs: {
          item: {
            newTables: [{ name: "Item", columns: [{ header: "Note" }] }],
          },
        },
      },
    });

    expect(service.batchUpdateCount()).toBe(0);
    expect(files?.alterations).toBeUndefined();
    expect(result).toMatchObject({
      outcome: { kind: "stopped" },
      stops: [
        {
          kind: "tableNameTaken",
          name: "Item",
          existing: "Item",
          tableId: "item",
        },
      ],
    });
  });

  it("stops before sending on a new tab whose title a live tab gives, naming it", () => {
    const service = seedFixture();

    const { result, files } = runItem({
      isSending: true,
      alterations: { newTabs: [{ title: "Item", tables: [paymentLog] }] },
    });

    expect(service.batchUpdateCount()).toBe(0);
    expect(files?.alterations).toBeUndefined();
    expect(service.grid.sheetTitles()).toEqual(["Item"]);
    expect(result).toMatchObject({
      outcome: { kind: "stopped" },
      stops: [
        {
          kind: "tabTitleTaken",
          title: "Item",
          existing: "Item",
          gid: itemGid,
        },
      ],
    });
  });

  describe("a new Table's rules", () => {
    const overThree = {
      columns: ["amount"],
      condition: { type: "NUMBER_GREATER" as const, values: ["3"] },
      format: { foregroundColor: "#0000ff" },
    };
    const withRule = {
      ...paymentLog,
      conditionalFormats: { add: [overThree] },
    };

    function overThreeRule(
      sheetId: number,
      startColumnIndex: number,
    ): GoogleAppsScript.Sheets.Schema.ConditionalFormatRule {
      return {
        ranges: [
          {
            sheetId,
            startRowIndex: tableHeaderRowIndex + 1,
            endRowIndex: tableHeaderRowIndex + 2,
            startColumnIndex,
            endColumnIndex: startColumnIndex + 1,
          },
        ],
        booleanRule: {
          condition: {
            type: "NUMBER_GREATER",
            values: [{ userEnteredValue: "3" }],
          },
          format: {
            textFormat: { foregroundColor: { red: 0, green: 0, blue: 1 } },
          },
        },
      };
    }

    it("sends a rule over a Table on a new tab in the run that creates it", () => {
      const service = seedFixture();

      const { result, files } = runItem({
        isSending: true,
        alterations: { newTabs: [{ title: "Payments", tables: [withRule] }] },
      });

      const gid = tabGidOf(mapOf(files), "Payments");
      expect(service.grid.sheet(gid).conditionalFormats).toEqual([
        overThreeRule(gid, startTableColIndex + 1),
      ]);
      expect(result).toMatchObject({
        outcome: { kind: "sent" },
        refusals: [],
        didntLand: [],
      });
      expect(files?.conditionalFormatConfigs).toContain("amount | ");
      expect(files?.alterations).toContain("makePendingAlterations({})");
    });

    it("compiles a rule over a Table beside Item against where it lands, after Item's new column", () => {
      const service = seedFixture();
      const itemEnd = itemTableEnd(service);

      const { result } = runItem({
        isSending: true,
        alterations: {
          tables: { item: { newColumns: [{ header: "Late fee" }] } },
          tabs: { item: { newTables: [{ ...withRule, gap: 1 }] } },
        },
      });

      const newTableStart = itemEnd + 1 + 1;
      expect(service.grid.sheet(itemGid).tables[1]?.range).toMatchObject({
        startColumnIndex: newTableStart,
      });
      expect(service.grid.sheet(itemGid).conditionalFormats).toContainEqual(
        overThreeRule(itemGid, newTableStart + 1),
      );
      expect(result).toMatchObject({
        outcome: { kind: "sent" },
        didntLand: [],
      });
    });

    it("refuses before any read a rule naming a column none of the new Table's headers gives", () => {
      const service = seedFixture();

      const { result } = runItem({
        isSending: true,
        alterations: {
          newTabs: [
            {
              title: "Payments",
              tables: [
                {
                  ...paymentLog,
                  conditionalFormats: {
                    add: [{ ...overThree, columns: ["amounts"] }],
                  },
                },
              ],
            },
          ],
        },
      });

      expect(service.batchUpdateCount()).toBe(0);
      expect(result.refusals).toEqual([
        {
          kind: "unknownColumn",
          table: "Payments",
          path: "newTabs[0].tables[0].conditionalFormats.add[0].columns",
          column: "amounts",
          columns: ["payee", "amount"],
        },
      ]);
    });

    it("refuses before any read a rule naming a column two of the new Table's headers give", () => {
      const service = seedFixture();

      const { result } = runItem({
        isSending: true,
        alterations: {
          tabs: {
            item: {
              newTables: [
                {
                  ...withRule,
                  columns: [{ header: "Amount" }, { header: "amount" }],
                },
              ],
            },
          },
        },
      });

      expect(service.batchUpdateCount()).toBe(0);
      expect(result.refusals).toEqual([
        {
          kind: "ambiguousColumn",
          table: "item",
          path: "newTables[0].conditionalFormats.add[0].columns",
          column: "amount",
          headers: ["Amount", "amount"],
        },
      ]);
    });

    it("throws before any send on a rule whose colour isn't six-digit hex", () => {
      const service = seedFixture();
      const blueByName = { ...overThree, format: { foregroundColor: "blue" } };

      expect(() =>
        runItem({
          isSending: true,
          alterations: {
            newTabs: [
              {
                title: "Payments",
                tables: [
                  { ...paymentLog, conditionalFormats: { add: [blueByName] } },
                ],
              },
            ],
          },
        }),
      ).toThrow(/blue/);
      expect(service.batchUpdateCount()).toBe(0);
    });
  });

  it("keeps Item's columns as they were", () => {
    seedFixture();

    const { files } = runItem({
      isSending: true,
      alterations: { tabs: { item: { newTables: [paymentLog] } } },
    });

    expect(mapOf(files).tables.item?.columns).toEqual(
      Object.fromEntries(
        Object.entries(itemColumns).map(([name, { columnId }]) => [
          name,
          columnId,
        ]),
      ),
    );
  });
});
