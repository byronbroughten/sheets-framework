import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import type { IdNameMapGeneric } from "../01_SpreadsheetSchema/configReaders/idNameMapTypes.ts";
import type { PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState.ts";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals.ts";
import { parseIdNameMapFileSource } from "./parseIdNameMapFileSource.ts";
import {
  emittedItemColumn,
  runItem,
  seedFixture,
} from "./syncCoordinatorTestSupport.ts";

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("SyncCoordinator.run a new Table's formulas and column types", () => {
  const payeeRule = { condition: "ONE_OF_LIST", values: ["Alice", "Bob"] };
  // A formula naming each other new column by the name its header gives, "Payee Name" as payeeName.
  const payments: PendingAlterations = {
    newTabs: [
      {
        title: "Ledger",
        tables: [
          {
            name: "Payments",
            columns: [
              {
                header: "Payee Name",
                columnType: { type: "DROPDOWN", rule: payeeRule },
              },
              { header: "Amount", emptyValueAllowed: true },
              { header: "Doubled", formula: "=${@amount}*2" },
              { header: "Label", formula: "=${@payeeName}" },
            ],
          },
        ],
      },
    ],
  };

  function ledgerGidOf(map: IdNameMapGeneric): number {
    const found = Object.entries(map.tabs).find(
      ([, tab]) => tab.title === "Ledger",
    );
    return Number(Val.assert(found, "mapped tab Ledger")[0]);
  }

  it("creates the Table with its dropdown and fills each formula column in one run, and the configs record both in brace form", () => {
    const service = seedFixture();

    const { result, files } = runItem({
      isSending: true,
      alterations: payments,
    });

    const emitted = Val.assert(files, "emitted files");
    const map = parseIdNameMapFileSource(emitted.idNameMap);
    const sheet = service.grid.sheet(ledgerGidOf(map));
    const [table] = sheet.tables;
    expect(table?.name).toBe("Payments");
    expect(table?.columnProperties?.[0]).toMatchObject({
      columnName: "Payee Name",
      columnType: "DROPDOWN",
      dataValidationRule: {
        condition: {
          type: "ONE_OF_LIST",
          values: [{ userEnteredValue: "Alice" }, { userEnteredValue: "Bob" }],
        },
      },
    });
    expect(sheet.bodyRows().map((row) => row.slice(2))).toEqual([
      [
        { isFormula: true, value: "=SINGLE(Payments[Amount])*2" },
        { isFormula: true, value: "=SINGLE(Payments[Payee Name])" },
      ],
    ]);
    expect(emittedItemColumn(emitted, "doubled")).toContain(
      '"formula": "=${@amount}*2"',
    );
    expect(emittedItemColumn(emitted, "label")).toContain(
      '"formula": "=${@payeeName}"',
    );
    expect(emittedItemColumn(emitted, "payeeName")).toContain(
      '"columnType": {"type":"DROPDOWN","rule":{"condition":"ONE_OF_LIST","values":["Alice","Bob"]}}',
    );
    expect(
      Object.values(map.tables).find(({ name }) => name === "payments")
        ?.columns,
    ).toMatchObject({
      payeeName: expect.any(String),
      doubled: expect.any(String),
    });
    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      landed: [
        { kind: "send", subject: { tab: "Ledger", path: "newTabs" } },
        { kind: "createdTable", table: "payments", tab: "Ledger" },
      ],
    });
  });

  it("refuses a brace naming no planned column before any read", () => {
    const service = seedFixture();

    const { result } = runItem({
      isSending: true,
      alterations: {
        newTabs: [
          {
            title: "Ledger",
            tables: [
              {
                name: "Payments",
                columns: [
                  { header: "Amount" },
                  { header: "Doubled", formula: "=${@amountDue}*2" },
                ],
              },
            ],
          },
        ],
      },
    });

    expect(result.refusals).toEqual([
      expect.objectContaining({ kind: "unknownBrace", brace: "${@amountDue}" }),
    ]);
    expect(service.getCalls).toEqual([]);
  });
});
