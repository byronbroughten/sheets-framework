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
