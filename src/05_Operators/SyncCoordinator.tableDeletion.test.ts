import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { type PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState";
import { getTableTraitByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  buildGridRows,
  type FakeSheetsService,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { fakeSpreadsheet } from "../testSupport/fakeSheetsService/fakeSpreadsheet";
import { parseIdNameMapFileSource } from "./parseIdNameMapFileSource";
import { type ConfigRegeneration, SyncCoordinator } from "./SyncCoordinator";
import {
  bodyRowsWithNote,
  frame,
  itemAtEachFlush,
  type ItemAtFlush,
  type ItemAtFlushProps,
  itemColumns,
  itemGid,
  itemSheet,
  notesGid,
  notesSheet,
  previousItemMap,
  runItem,
  seedFixture,
} from "./syncCoordinatorTestSupport";

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("SyncCoordinator.run Table deletion", () => {
  const deleteItem: PendingAlterations = {
    tables: { item: { deleted: true } },
  };
  const itemSubject = { table: "item", path: "deleted" };
  const itemPrefix = getTableTraitByName("item", "idPrefix");
  const everyItemColIndex = Object.keys(itemColumns).map(
    (_, colIndex) => colIndex,
  );

  function retiredPrefixesOf(files: ConfigRegeneration): unknown {
    const match = files.retiredIdPrefixes.match(
      /makeRetiredIdPrefixes\((.*)\);/,
    );
    return JSON.parse(Val.assert(match?.[1], "emitted retired prefixes"));
  }

  it("plans a deletion on a dry run, sending nothing", () => {
    const service = seedFixture();

    const { result } = runItem({ isSending: false, alterations: deleteItem });

    expect(service.batchUpdateCount()).toBe(0);
    expect(result).toMatchObject({
      mode: "dry run",
      outcome: { kind: "ready" },
      refusals: [],
      plan: [{ subject: itemSubject, what: "delete the Table" }],
    });
  });

  it("deletes the Table on a send, leaving its tab, and drops it from the configs, the ID-name map and alterations.ts, retiring its prefix", () => {
    const service = seedFixture({ item: { bodyRows: bodyRowsWithNote } });

    const { result, files } = runItem({
      isSending: true,
      alterations: deleteItem,
    });

    const emitted = Val.assert(files, "emitted files");
    const sheet = service.grid.sheet(itemGid);
    expect(sheet.tables).toEqual([]);
    expect(sheet.title).toBe("Item");
    expect(
      sheet
        .values()
        .flat()
        .filter((cell) => cell !== "" && cell !== null),
    ).toEqual([]);
    expect(emitted.tableConfigs).not.toContain('"item"');
    expect(emitted.columnConfigs).not.toContain('"item"');
    expect(
      parseIdNameMapFileSource(emitted.idNameMap).tables,
    ).not.toHaveProperty("item");
    expect(retiredPrefixesOf(emitted)).toContain(itemPrefix);
    expect(emitted.alterations).toContain("makePendingAlterations({})");
    expect(result).toMatchObject({
      mode: "send",
      outcome: { kind: "sent" },
      landed: [{ kind: "send", subject: itemSubject }],
      didntLand: [],
    });
  });

  it("reports a deletion the refetch still shows the Table for under Didn't land, the entry staying pending", () => {
    stubSheetsService({
      sheets: [itemSheet()],
      afterEachBatchUpdate(spreadsheet) {
        spreadsheet.sheets = fakeSpreadsheet.init([itemSheet()]).sheets;
      },
    });

    const { result, files } = runItem({
      isSending: true,
      alterations: deleteItem,
    });

    const emitted = Val.assert(files, "emitted files");
    expect(emitted.tableConfigs).toContain('"item"');
    expect(emitted.alterations).toBeUndefined();
    expect(result).toMatchObject({
      outcome: { kind: "sent", summary: "nothing landed; 1 didn't land" },
      landed: [],
      didntLand: [
        { subject: itemSubject, what: "delete the Table; Sheets kept false" },
      ],
    });
  });

  function itemAtFirstFlush(
    props: Pick<ItemAtFlushProps, "isSending" | "isHostWithholding">,
  ): ItemAtFlush | undefined {
    return itemAtEachFlush({ ...props, alterations: deleteItem })[0];
  }

  it("reads every column of the body before the send's flush, holding the cell a person entered", () => {
    expect(itemAtFirstFlush({ isSending: true })).toEqual({
      colIndexesRead: everyItemColIndex,
      note: "kept",
    });
  });

  it("reads every column of the body before a withheld dry run's flush", () => {
    expect(
      itemAtFirstFlush({ isSending: false, isHostWithholding: true }),
    ).toEqual({ colIndexesRead: everyItemColIndex, note: "kept" });
  });

  it("refuses a deletion of a Table without API access as noApiAccess", () => {
    seedFixture({ extraSheets: [notesSheet()] });

    const { result } = SyncCoordinator.init().run({
      isSending: true,
      intendedState: {
        alterations: {
          tables: { deskNotes: { deleted: true } },
        } as PendingAlterations,
      },
      previousIdNameMap: {
        tables: {
          ...previousItemMap().tables,
          notes: { name: "deskNotes", sheetGid: notesGid },
        },
        tabs: previousItemMap().tabs,
      },
      forcedTraits: [],
      frame,
    });

    expect(result).toMatchObject({
      outcome: { kind: "refused" },
      refusals: [{ kind: "noApiAccess", table: "deskNotes", path: "deleted" }],
    });
  });

  describe("a Table a person deleted", () => {
    function seedWithItemGone(): FakeSheetsService {
      return stubSheetsService({
        sheets: [{ sheetId: itemGid, title: "Item", rows: buildGridRows({}) }],
      });
    }

    it("leaves the configs with no conflict, retiring its prefix and naming it in a notice", () => {
      seedWithItemGone();

      const { result, files } = runItem({ isSending: true });

      const emitted = Val.assert(files, "emitted files");
      expect(emitted.tableConfigs).not.toContain('"item"');
      expect(retiredPrefixesOf(emitted)).toContain(itemPrefix);
      expect(result).toMatchObject({
        outcome: { kind: "sent" },
        stops: [],
        notices: [{ kind: "tableGone", table: "item" }],
      });
    });

    it("plans nothing for an entry deleting it, and the next run refuses that entry", () => {
      seedWithItemGone();

      const first = runItem({ isSending: true, alterations: deleteItem });
      const next = SyncCoordinator.init().run({
        isSending: true,
        intendedState: { alterations: deleteItem },
        previousIdNameMap: parseIdNameMapFileSource(
          Val.assert(first.files, "first run's files").idNameMap,
        ),
        forcedTraits: [],
        frame,
      });

      expect(first.result).toMatchObject({ stops: [], plan: [] });
      expect(next.result).toMatchObject({
        outcome: { kind: "refused" },
        refusals: [{ kind: "noSuchTable", entry: "tables", name: "item" }],
      });
    });
  });
});
