import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { type PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState";
import { getTableTraitByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  buildGridRows,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { SyncCoordinator } from "./SyncCoordinator";
import {
  frame,
  itemGid,
  previousItemMap,
  retiredPrefixesOf,
  runItem,
  seedFixture,
} from "./syncCoordinatorTestSupport";

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("SyncCoordinator.run tab deletion", () => {
  describe("a Table or column deletion inside the tab being deleted", () => {
    it.each<{
      row: string;
      alterations: PendingAlterations;
      touchedBy: string;
    }>([
      {
        row: "the tab's Table",
        alterations: {
          tabs: { item: { deleted: true } },
          tables: { item: { deleted: true } },
        },
        touchedBy: "tables.item",
      },
      {
        row: "a column of the tab's Table",
        alterations: {
          tabs: { item: { deleted: true } },
          tables: { item: { columns: { optionalNote: { deleted: true } } } },
        },
        touchedBy: "tables.item.columns.optionalNote",
      },
    ])(
      "refuses $row before any read, naming both entries and the Table the tab's deletion takes",
      ({ alterations, touchedBy }) => {
        const service = seedFixture();

        const { result } = runItem({ isSending: true, alterations });

        expect(service.getByDataFilterCalls).toEqual([]);
        expect(service.batchUpdateCount()).toBe(0);
        expect(result.outcome).toMatchObject({ kind: "refused" });
        expect(result.refusals).toContainEqual({
          kind: "touchesDeletion",
          deleted: { tab: "item" },
          deletedAt: "tabs.item",
          touchedBy,
          takenTable: "item",
        });
      },
    );
  });

  describe("a tab a person deleted", () => {
    const otherGid = 999002;

    function seedWithItemTabGone(): void {
      stubSheetsService({
        sheets: [
          { sheetId: otherGid, title: "Other", rows: buildGridRows({}) },
        ],
      });
    }

    it("drops the tab and its Tables from the configs with no conflict, retiring their prefixes and naming the tab beside its Tables", () => {
      seedWithItemTabGone();

      const { result, files } = runItem({ isSending: true });

      const emitted = Val.assert(files, "emitted files");
      expect(emitted.tabConfigs).not.toContain('"item"');
      expect(emitted.tableConfigs).not.toContain('"item"');
      expect(retiredPrefixesOf(emitted)).toContain(
        getTableTraitByName("item", "idPrefix"),
      );
      expect(result).toMatchObject({
        outcome: { kind: "sent" },
        stops: [],
        notices: [
          { kind: "tabGone", tab: "item" },
          { kind: "tableGone", table: "item" },
        ],
      });
    });

    it("names each gone tab directly before its own Tables' lines", () => {
      seedWithItemTabGone();
      const previous = previousItemMap();
      // GID keys list in numeric order, so this tab's notice comes after Item's.
      const logGid = itemGid + 1;

      const { result } = SyncCoordinator.init().run({
        isSending: true,
        intendedState: { alterations: {} },
        previousIdNameMap: {
          tables: previous.tables,
          tabs: {
            ...previous.tabs,
            [String(logGid)]: { name: "log", title: "Log", tableIds: [] },
          },
        },
        forcedTraits: [],
        frame,
      });

      expect(result.notices).toEqual([
        { kind: "tabGone", tab: "item" },
        { kind: "tableGone", table: "item" },
        { kind: "tabGone", tab: "log" },
      ]);
    });
  });
});
