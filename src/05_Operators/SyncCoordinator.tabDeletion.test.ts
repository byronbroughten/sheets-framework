import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { type IdNameMapGeneric } from "../01_SpreadsheetSchema/configReaders/idNameMapTypes";
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
import { SyncCoordinator, type SyncRun } from "./SyncCoordinator";
import {
  bodyRowsWithNote,
  columnConfigs,
  dryRunBesideLog,
  frame,
  itemGid,
  itemSheet,
  logFormulaCell,
  previousItemMap,
  retiredPrefixesOf,
  runItem,
  seedFixture,
  uncheckedOutsideFormulasNotice,
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

// What the item tab held, and whether the run had read it whole, as a flush began.
interface ItemTabAtFlush {
  readsWholeTab: boolean;
  cells: unknown[];
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
      "refuses $row before any read, once, naming both entries and the Table the tab's deletion takes",
      ({ alterations, touchedBy }) => {
        const service = seedFixture();

        const { result } = runItem({ isSending: true, alterations });

        expect(service.getByDataFilterCalls).toEqual([]);
        expect(service.batchUpdateCount()).toBe(0);
        expect(result.outcome).toMatchObject({ kind: "refused" });
        expect(result.refusals).toEqual([
          {
            kind: "touchesDeletion",
            deleted: { tab: "item" },
            deletedAt: "tabs.item",
            touchedBy,
            takenTable: "item",
          },
        ]);
      },
    );

    it("refuses a changed column of the tab's Table once, naming the column entry", () => {
      seedFixture();

      const { result } = runItem({
        isSending: true,
        alterations: {
          tabs: { item: { deleted: true } },
          tables: { item: { columns: { optionalNote: { header: "Memo" } } } },
        },
      });

      expect(result.refusals).toEqual([
        {
          kind: "touchesDeletion",
          deleted: { tab: "item" },
          deletedAt: "tabs.item",
          touchedBy: "tables.item.columns.optionalNote",
        },
      ]);
    });
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

  describe("its Tables in the previous ID-name map", () => {
    function dryRunDeletingItemTab(
      previousIdNameMap: IdNameMapGeneric,
    ): SyncRun {
      seedFixture();
      return SyncCoordinator.init().run({
        isSending: false,
        intendedState: { alterations: deleteItemTab },
        previousIdNameMap,
        forcedTraits: [],
        frame,
      });
    }

    it("throws on a tab naming a Table the map lacks, since one build writes both", () => {
      const previous = previousItemMap();
      const itemTab = Val.assert(previous.tabs[String(itemGid)], "item tab");
      itemTab.tableIds.push("missing");

      expect(() => dryRunDeletingItemTab(previous)).toThrow(
        "the ID-name map's Table missing",
      );
    });

    it("plans the deletion of a tab holding a Table without API access, which the Table configs lack", () => {
      const previous = previousItemMap();
      previous.tables.notes = { name: "notes", sheetGid: itemGid };
      const itemTab = Val.assert(previous.tabs[String(itemGid)], "item tab");
      itemTab.tableIds.push("notes");

      const { result } = dryRunDeletingItemTab(previous);

      expect(result.refusals).toEqual([]);
      expect(result.plan).toEqual([
        { subject: tabSubject, what: "delete the tab", effects: [] },
      ]);
    });
  });

  describe("its outside references", () => {
    const nameHeader = columnConfigs.item.name.header;

    it.each<{ row: string; logFormula: string; formula: string }>([
      {
        row: "a reference to its Table",
        logFormula: `=COUNTA(item[${nameHeader}])`,
        formula: "=COUNTA(${item[name]})",
      },
      {
        row: "a reference to a header its Table's configs don't hold",
        logFormula: "=COUNTA(item[No Such Header])",
        formula: "=COUNTA(item[No Such Header])",
      },
      {
        row: "its title as a quoted sheet reference",
        logFormula: "='Item'!B2",
        formula: "='Item'!B2",
      },
      {
        row: "its title as a bare sheet reference",
        logFormula: "=SUM(Item!B2:B9)",
        formula: "=SUM(Item!B2:B9)",
      },
    ])(
      "lists a formula column on another tab holding $row under Check before sending, and counts it on the plan line",
      ({ logFormula, formula }) => {
        const { result } = dryRunBesideLog({
          alterations: deleteItemTab,
          logFormula,
        });

        expect(result.checks).toEqual([
          {
            kind: "outsideReferences",
            subject: { tab: "item" },
            effect: "ref",
            references: [{ cell: logFormulaCell, formula }],
          },
        ]);
        expect(result.plan).toEqual([
          {
            subject: tabSubject,
            what: "delete the tab",
            effects: [{ kind: "outsideReferences", count: 1, effect: "ref" }],
          },
        ]);
      },
    );

    it("leaves out a referencing formula in a Table deleted in the same run", () => {
      const { result } = dryRunBesideLog({
        alterations: { ...deleteItemTab, tables: { log: { deleted: true } } },
        logFormula: "='Item'!B2",
      });

      expect(result.checks).toEqual([]);
      expect(result.plan.map(({ effects }) => effects)).toEqual([[], []]);
    });

    it("says in a notice that formulas outside the API-access Tables' formula columns aren't checked", () => {
      const { result } = dryRunBesideLog({
        alterations: deleteItemTab,
        logFormula: "=1",
      });

      expect(result.checks).toEqual([]);
      expect(result.notices).toContainEqual(uncheckedOutsideFormulasNotice);
    });
  });

  describe("before the flush", () => {
    function atFirstFlush(props: {
      isSending: boolean;
      isHostWithholding?: boolean;
    }): ItemTabAtFlush | undefined {
      const atFlush: ItemTabAtFlush[] = [];
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
