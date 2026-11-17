import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { type PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState";
import { getTableTraitByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  buildGridRows,
  type FakeCell,
  type FakeSheetProperties,
  type FakeSheetsService,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { fakeSpreadsheet } from "../testSupport/fakeSheetsService/fakeSpreadsheet";
import { parseIdNameMapFileSource } from "./parseIdNameMapFileSource";
import { SyncCoordinator } from "./SyncCoordinator";
import {
  bodyRowsWithNote,
  frame,
  itemGid,
  itemSheet,
  previousItemMap,
  retiredPrefixesOf,
  runItem,
} from "./syncCoordinatorTestSupport";

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

// A tab left on the sheet, since Google never deletes a spreadsheet's last one.
const otherSheet: FakeSheetProperties = {
  sheetId: 424242,
  title: "Other",
  rows: buildGridRows({}),
};

const looseCell = "loose";

// Item with a person's note in the body and a cell below its Table, outside any Table.
function itemSheetWithLooseCell(): FakeSheetProperties {
  const sheet = itemSheet({ bodyRows: bodyRowsWithNote });
  const table = Val.assert(sheet.tables?.[0], "the Item Table");
  const looseRowIndex = Val.assert(table.endRowIndex, "end row") + 2;
  const rows: FakeCell[][] = (sheet.rows ?? []).map((row) => [...row]);
  while (rows.length <= looseRowIndex) rows.push([]);
  rows[looseRowIndex] = [looseCell];
  return { ...sheet, rows };
}

interface ReadGridRange {
  sheetId?: number;
  startRowIndex?: number;
  startColumnIndex?: number;
}

function readsWholeItemTab(service: FakeSheetsService): boolean {
  return service.getByDataFilterCalls.some((resource) => {
    const { dataFilters = [] } = resource as {
      dataFilters?: { gridRange?: ReadGridRange }[];
    };
    return dataFilters.some(
      ({ gridRange }) =>
        gridRange?.sheetId === itemGid &&
        gridRange.startRowIndex === undefined &&
        gridRange.startColumnIndex === undefined,
    );
  });
}

describe("SyncCoordinator.run tab deletion", () => {
  const deleteItemTab: PendingAlterations = {
    tabs: { item: { deleted: true } },
  };
  const tabSubject = { tab: "item", path: "deleted" };
  const itemPrefix = getTableTraitByName("item", "idPrefix");

  it("plans the deletion on a dry run, sending nothing", () => {
    const service = stubSheetsService({ sheets: [itemSheet(), otherSheet] });

    const { result } = runItem({
      isSending: false,
      alterations: deleteItemTab,
    });

    expect(service.batchUpdateCount()).toBe(0);
    expect(result).toMatchObject({
      mode: "dry run",
      outcome: { kind: "ready" },
      refusals: [],
      plan: [{ subject: tabSubject, what: "delete the tab" }],
    });
  });

  it("deletes the tab on a send, and drops it and its Table from the configs, the ID-name map and alterations.ts, retiring the Table's prefix", () => {
    const service = stubSheetsService({ sheets: [itemSheet(), otherSheet] });

    const { result, files } = runItem({
      isSending: true,
      alterations: deleteItemTab,
    });

    const emitted = Val.assert(files, "emitted files");
    expect(service.grid.sheetTitles()).toEqual(["Other"]);
    expect(emitted.tableConfigs).not.toContain('"item"');
    expect(emitted.columnConfigs).not.toContain('"item"');
    expect(emitted.tabConfigs).not.toContain('"item"');
    const idNameMap = parseIdNameMapFileSource(emitted.idNameMap);
    expect(idNameMap.tables).not.toHaveProperty("item");
    expect(idNameMap.tabs).not.toHaveProperty(String(itemGid));
    expect(retiredPrefixesOf(emitted)).toContain(itemPrefix);
    expect(emitted.alterations).toContain("makePendingAlterations({})");
    expect(result).toMatchObject({
      mode: "send",
      outcome: { kind: "sent" },
      landed: [{ kind: "send", subject: tabSubject }],
      didntLand: [],
    });
  });

  it("reports a deletion the refetch still shows the tab for under Didn't land, the entry staying pending", () => {
    stubSheetsService({
      sheets: [itemSheet(), otherSheet],
      afterEachBatchUpdate(spreadsheet) {
        spreadsheet.sheets = fakeSpreadsheet.init([
          itemSheet(),
          otherSheet,
        ]).sheets;
      },
    });

    const { result, files } = runItem({
      isSending: true,
      alterations: deleteItemTab,
    });

    const emitted = Val.assert(files, "emitted files");
    expect(emitted.tableConfigs).toContain('"item"');
    expect(emitted.alterations).toBeUndefined();
    expect(result).toMatchObject({
      outcome: { kind: "sent", summary: "nothing landed; 1 didn't land" },
      landed: [],
      didntLand: [
        { subject: tabSubject, what: "delete the tab; Sheets kept false" },
      ],
    });
  });

  describe("before the flush", () => {
    function atFirstFlush(props: {
      isSending: boolean;
      isHostWithholding?: boolean;
    }): { readsWholeTab: boolean; cells: unknown[] } | undefined {
      const atFlush: { readsWholeTab: boolean; cells: unknown[] }[] = [];
      const service = stubSheetsService({
        sheets: [itemSheetWithLooseCell(), otherSheet],
        isDryRun: !props.isSending,
        beforeEachBatchUpdate: () => {
          atFlush.push({
            readsWholeTab: readsWholeItemTab(service),
            cells: service.grid.sheet(itemGid).values().flat(),
          });
        },
      });
      SyncCoordinator.init().run({
        ...props,
        intendedState: { alterations: deleteItemTab },
        previousIdNameMap: previousItemMap(),
        forcedTraits: [],
        frame,
      });
      return atFlush[0];
    }

    it("reads the whole tab before the send's flush, a Table's body data and a cell outside any Table among it", () => {
      const first = Val.assert(atFirstFlush({ isSending: true }), "a flush");

      expect(first.readsWholeTab).toBe(true);
      expect(first.cells).toEqual(expect.arrayContaining(["kept", looseCell]));
    });

    it("reads the whole tab before a withheld dry run's flush", () => {
      const first = Val.assert(
        atFirstFlush({ isSending: false, isHostWithholding: true }),
        "a flush",
      );

      expect(first.readsWholeTab).toBe(true);
    });
  });
});
