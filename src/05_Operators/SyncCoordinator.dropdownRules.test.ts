import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { getTableTraitByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes.ts";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals.ts";
import { fakeTableSheet } from "../testSupport/fakeSheetsService/fakeTableSheet.ts";
import { SyncCoordinator } from "./SyncCoordinator.ts";
import {
  columnsByHeader,
  itemColumns,
  runWithNoAlterations,
  seedFixture,
} from "./syncCoordinatorTestSupport.ts";

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("SyncCoordinator dropdown rules", () => {
  const ledgerGid = 889102;
  const idHeader = itemColumns.id.header;

  // A ledger whose Status dropdown takes its values from the Item Table's ID column, named by the Table's live name.
  function ledgerNaming(itemTableName: string) {
    return fakeTableSheet.build({
      sheetId: ledgerGid,
      title: "Ledger",
      tableId: getTableTraitByName("runItem", "tableId"),
      name: "Ledger",
      columnConfigs: columnsByHeader(["Status"]),
      columnNames: ["Status"],
      bodyRows: [{ Status: "r:itm:fixture" }],
      columnTypes: { 0: "DROPDOWN" },
      columnValidationConditionTypes: { 0: "ONE_OF_RANGE" },
      columnValidationValues: { 0: [`=${itemTableName}[${idHeader}]`] },
    });
  }

  function liveStatusColumnType(itemTableName: string): unknown {
    seedFixture({
      item: { name: itemTableName },
      extraSheets: [ledgerNaming(itemTableName)],
    });
    const coordinator = SyncCoordinator.init();
    runWithNoAlterations(coordinator);
    return coordinator.columnConfigOperator.traitReadings().live.tables.runItem
      ?.columns.status?.columnType;
  }

  it("reads a rule naming a renamed Table as unchanged, in brace form", () => {
    const expected = {
      type: "DROPDOWN",
      rule: { condition: "ONE_OF_RANGE", values: ["=${item[id]}"] },
    };

    const configsName = getTableTraitByName("item", "tableName");

    expect(liveStatusColumnType(configsName)).toEqual(expected);
    expect(liveStatusColumnType("Stock")).toEqual(expected);
  });
});
