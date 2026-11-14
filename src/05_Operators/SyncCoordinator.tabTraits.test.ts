import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { type PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  type FakeSheetProperties,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { parseIdNameMapFileSource } from "./parseIdNameMapFileSource";
import { SyncCoordinator, type SyncRun } from "./SyncCoordinator";
import {
  frame,
  itemGid,
  itemSheet,
  previousItemMap,
  seedFixture,
} from "./syncCoordinatorTestSupport";

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("SyncCoordinator.run tab traits", () => {
  // The dev Test tab, with no Table, in the previous map beside Item.
  const testTabGid = 604825146;
  const testTabSheet: FakeSheetProperties = {
    sheetId: testTabGid,
    title: "Test",
  };

  const layoutTabSheet: FakeSheetProperties = {
    sheetId: 604825147,
    title: "Layout",
  };

  function runTabs(
    alterations: PendingAlterations,
    extraTabs: FakeSheetProperties[] = [testTabSheet],
  ): SyncRun {
    const previousIdNameMap = previousItemMap();
    for (const { sheetId, title } of extraTabs) {
      previousIdNameMap.tabs[String(sheetId)] = {
        name: title.toLowerCase(),
        title,
        tableIds: [],
      };
    }
    return SyncCoordinator.init().run({
      isSending: true,
      intendedState: { alterations },
      previousIdNameMap,
      forcedTraits: [],
      frame,
    });
  }

  it("sets a tab's colour, frozen counts and hidden, and the configs record them", () => {
    const service = seedFixture();

    const { result, files } = runTabs({
      tabs: {
        item: {
          color: "#ff0000",
          frozenRows: 2,
          frozenColumns: 1,
          hidden: true,
        },
      },
    });

    expect(service.grid.sheet(itemGid).tabProperties).toEqual({
      hidden: true,
      tabColorStyle: { rgbColor: { red: 1, green: 0, blue: 0 } },
      frozenRowCount: 2,
      frozenColumnCount: 1,
    });
    expect(Val.assert(files, "emitted files").tabConfigs).toContain(
      '"item": { "sheetGid": 1100001, "title": "Item", "color": "#ff0000", "hidden": true, "frozenRows": 2, "frozenColumns": 1 }',
    );
    expect(result).toMatchObject({
      outcome: { kind: "sent", summary: "4 alterations landed" },
    });
  });

  it("retitles a tab, and the ID-name map follows the new title", () => {
    const service = seedFixture();

    const { files } = runTabs({ tabs: { item: { title: "Stock" } } });

    expect(service.grid.sheet(itemGid).title).toBe("Stock");
    expect(
      parseIdNameMapFileSource(Val.assert(files, "emitted files").idNameMap)
        .tabs[String(itemGid)],
    ).toMatchObject({ name: "stock", title: "Stock" });
  });

  it("moves a tab before another", () => {
    const service = seedFixture({ extraSheets: [testTabSheet] });

    const { result } = runTabs({ tabs: { test: { before: "item" } } });

    expect(service.grid.sheetTitles()).toEqual(["Test", "Item"]);
    expect(result).toMatchObject({
      landed: [{ subject: { tab: "test", path: "before" } }],
    });
  });

  it("moves two tabs in one run, the second counted from where the first left the tabs", () => {
    const extraTabs = [testTabSheet, layoutTabSheet];
    const service = seedFixture({ extraSheets: extraTabs });

    runTabs(
      { tabs: { layout: { before: "item" }, test: { before: "item" } } },
      extraTabs,
    );

    expect(service.grid.sheetTitles()).toEqual(["Layout", "Test", "Item"]);
  });

  it("skips a before the live order already satisfies", () => {
    const service = seedFixture({ extraSheets: [testTabSheet] });

    const { result } = runTabs({ tabs: { item: { before: "test" } } });

    expect(service.batchUpdateCount()).toBe(0);
    expect(service.grid.sheetTitles()).toEqual(["Item", "Test"]);
    expect(result).toMatchObject({ landed: [] });
  });

  it("stops on a person's tab edit since the last sync, sending nothing", () => {
    const service = stubSheetsService({
      sheets: [{ ...itemSheet(), frozenRowCount: 3 }],
    });

    const { result } = runTabs({ tabs: { item: { frozenRows: 2 } } });

    expect(service.batchUpdateCount()).toBe(0);
    expect(service.grid.sheet(itemGid).tabProperties.frozenRowCount).toBe(3);
    expect(result).toMatchObject({
      stops: [
        {
          subject: { tab: "item", path: "frozenRows" },
          label: "both changed",
          agreed: "0",
          live: "3",
        },
      ],
    });
  });
});
