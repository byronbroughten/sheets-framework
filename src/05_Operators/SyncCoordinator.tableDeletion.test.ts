import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { type IdNameMapGeneric } from "../01_SpreadsheetSchema/configReaders/idNameMapTypes";
import { type PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState";
import { getTableTraitByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  buildGridRows,
  type FakeSheetProperties,
  type FakeSheetsService,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { a1Notation } from "../testSupport/fakeSheetsService/a1Notation";
import { fakeSpreadsheet } from "../testSupport/fakeSheetsService/fakeSpreadsheet";
import { fakeTableSheet } from "../testSupport/fakeSheetsService/fakeTableSheet";
import { parseIdNameMapFileSource } from "./parseIdNameMapFileSource";
import { SyncCoordinator, type SyncRun } from "./SyncCoordinator";
import {
  bodyRowsWithNote,
  columnConfigs,
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

describe("SyncCoordinator.run Table deletion", () => {
  const deleteItem: PendingAlterations = {
    tables: { item: { deleted: true } },
  };
  const itemSubject = { table: "item", path: "deleted" };
  const itemPrefix = getTableTraitByName("item", "idPrefix");
  const everyItemColIndex = Object.keys(itemColumns).map(
    (_, colIndex) => colIndex,
  );

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

  describe("its conditional-format rules", () => {
    const itemTable = Val.assert(itemSheet().tables?.[0], "the Item Table");
    const firstColIndex = Val.assert(itemTable.startColumnIndex, "first col");
    const headerRowIndex = Val.assert(itemTable.startRowIndex, "header row");
    const endRowIndex = Val.assert(itemTable.endRowIndex, "end row");
    const firstColumn = Val.assert(Object.keys(itemColumns)[0], "first column");

    function trueRule(
      endRow: number,
    ): GoogleAppsScript.Sheets.Schema.ConditionalFormatRule {
      return {
        ranges: [
          {
            sheetId: itemGid,
            startRowIndex: headerRowIndex + 1,
            endRowIndex: endRow,
            startColumnIndex: firstColIndex,
            endColumnIndex: firstColIndex + 1,
          },
        ],
        booleanRule: {
          condition: {
            type: "CUSTOM_FORMULA",
            values: [{ userEnteredValue: "=TRUE" }],
          },
          format: { backgroundColor: { red: 1, green: 0, blue: 0 } },
        },
      };
    }

    function seedWithRules(): FakeSheetsService {
      return stubSheetsService({
        sheets: [
          {
            ...itemSheet(),
            conditionalFormats: [
              trueRule(endRowIndex),
              trueRule(endRowIndex + 5),
            ],
          },
        ],
      });
    }

    it("plans the rule wholly over the Table as removed and reports the one running past it, which stays on the sheet", () => {
      const service = seedWithRules();

      const { result } = runItem({ isSending: true, alterations: deleteItem });

      expect(result.plan).toEqual([
        {
          subject: itemSubject,
          what: "delete the Table",
          effects: [
            {
              kind: "rulesRemoved",
              handles: [`${firstColumn} | =TRUE | background #ff0000`],
            },
            {
              kind: "rulePartlyOver",
              tab: "item",
              range: "A5:A10",
              formula: "=TRUE",
            },
          ],
        },
      ]);
      expect(
        service.grid
          .sheet(itemGid)
          .conditionalFormats?.map(({ ranges }) => ranges?.[0]?.endRowIndex),
      ).toEqual([endRowIndex + 5]);
    });
  });

  describe("its outside references", () => {
    const logGid = getTableTraitByName("log", "sheetGid");
    const logColumns = columnConfigs.log;
    const logColumnNames = Object.keys(logColumns);
    const formulaColumn = Val.assert(logColumnNames.at(-1), "a Log column");
    const formulaLetter = a1Notation.columnLetters(
      logColumnNames.indexOf(formulaColumn),
    );
    const nameHeader = columnConfigs.item.name.header;
    const uncheckedNotice = {
      kind: "text",
      text: "A Table deletion's outside formulas are checked only in the API-access Tables' formula columns; a formula anywhere else that reads a deleted Table isn't listed.",
    };

    // The brace decompile matches the Sheets name the previous configs record, which for item is "item".
    function logSheet(): FakeSheetProperties {
      return fakeTableSheet.build({
        sheetId: logGid,
        title: "Log",
        tableId: getTableTraitByName("log", "tableId"),
        name: getTableTraitByName("log", "tableName"),
        columnConfigs: logColumns,
        columnNames: logColumnNames,
        bodyRows: [
          {
            [formulaColumn]: {
              value: 1,
              isFormula: true,
              formula: `=COUNTA(item[${nameHeader}])`,
            },
          },
        ],
      });
    }

    function previousItemAndLogMap(): IdNameMapGeneric {
      const previous = previousItemMap();
      return {
        tables: {
          ...previous.tables,
          log: {
            name: "log",
            sheetGid: logGid,
            columns: Object.fromEntries(
              Object.entries(logColumns).map(([columnName, { columnId }]) => [
                columnName,
                columnId,
              ]),
            ),
          },
        },
        tabs: {
          ...previous.tabs,
          [String(logGid)]: {
            name: "log",
            title: "Log",
            tableIds: [getTableTraitByName("log", "tableId")],
          },
        },
      };
    }

    function runDeleting(alterations: PendingAlterations): SyncRun {
      seedFixture({ extraSheets: [logSheet()] });
      return SyncCoordinator.init().run({
        isSending: false,
        intendedState: { alterations },
        previousIdNameMap: previousItemAndLogMap(),
        forcedTraits: [],
        frame,
      });
    }

    it("lists a formula in another API-access Table's formula column under Check before sending, and counts it on the plan item", () => {
      const { result } = runDeleting(deleteItem);

      expect(result.checks).toEqual([
        {
          kind: "outsideReferences",
          subject: { table: "item" },
          effect: "blank",
          references: [
            {
              cell: `Log!${formulaLetter}5:${formulaLetter}5`,
              formula: "=COUNTA(${item[name]})",
            },
          ],
        },
      ]);
      expect(result.plan).toEqual([
        {
          subject: itemSubject,
          what: "delete the Table",
          effects: [{ kind: "outsideReferences", count: 1, effect: "blank" }],
        },
      ]);
    });

    it("says in a notice that formulas outside the API-access Tables' formula columns aren't checked", () => {
      const { result } = runDeleting(deleteItem);

      expect(result.notices).toContainEqual(uncheckedNotice);
    });

    it("leaves out a referencing formula in a Table deleted in the same run", () => {
      const { result } = runDeleting({
        tables: { item: { deleted: true }, log: { deleted: true } },
      });

      expect(result.checks).toEqual([]);
      expect(result.plan.map(({ effects }) => effects)).toEqual([[], []]);
    });

    it("lists nothing and gives no notice when no Table is deleted", () => {
      const { result } = runDeleting({});

      expect(result.checks).toEqual([]);
      expect(result.notices).not.toContainEqual(uncheckedNotice);
    });
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
