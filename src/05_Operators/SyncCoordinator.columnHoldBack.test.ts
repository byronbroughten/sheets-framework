import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { type PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState.ts";
import {
  layoutBodyRows,
  layoutGid,
  layoutOrigins,
  layoutSheet,
} from "../02_SpreadsheetRaw/spreadsheetRawTestSupport.ts";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals.ts";
import {
  type FakeSheetsService,
  stubSheetsService,
} from "../testSupport/fakeSheetsService.ts";
import { SyncCoordinator } from "./SyncCoordinator.ts";
import {
  frame,
  previousLayoutMap,
  runLayout,
} from "./syncCoordinatorTestSupport.ts";

type SheetsRule = GoogleAppsScript.Sheets.Schema.ConditionalFormatRule;

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

const right = layoutOrigins.layoutRight;
const rightEndRowIndex =
  right.startRowIndex + 1 + layoutBodyRows.layoutRight.length;

// Over every column of Right from `startColumnIndex`; reading no column, since the fake can't replay a left neighbour's insert under an anchored reference.
function rightRule(startColumnIndex: number, columnCount: number): SheetsRule {
  return {
    ranges: [
      {
        sheetId: layoutGid,
        startRowIndex: right.startRowIndex + 1,
        endRowIndex: rightEndRowIndex,
        startColumnIndex,
        endColumnIndex: startColumnIndex + columnCount,
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

function seedWith(rules: SheetsRule[]): FakeSheetsService {
  return stubSheetsService({
    sheets: [{ ...layoutSheet(), conditionalFormats: rules }],
  });
}

function tabRules(service: FakeSheetsService): SheetsRule[] {
  return service.grid.sheet(layoutGid).conditionalFormats ?? [];
}

const insertOnLeft: PendingAlterations = {
  tables: {
    layoutLeft: {
      newColumns: [{ columns: [{ header: "Note" }], before: "amount" }],
    },
  },
};
const insertOnRight: PendingAlterations = {
  tables: {
    layoutRight: {
      newColumns: [{ columns: [{ header: "Memo" }], before: "amount" }],
    },
  },
};
const deleteOnRight: PendingAlterations = {
  tables: { layoutRight: { columns: { entry: { deleted: true } } } },
};
// Held, so never sent: moving a Table's last column to its front is unmeasured, and Right has two columns.
const moveOnRight: PendingAlterations = {
  tables: {
    layoutRight: { moves: [{ columns: ["amount"], before: "entry" }] },
  },
};
const heldBackWhat =
  "held back: layoutLeft's column change shifts layoutRight's columns this run, so it is sent next run";

describe("SyncCoordinator.run holds back a Table's own column changes a left neighbour's column change shifts", () => {
  it("sends the left Table's insert, holds back the right Table's insert and its rule rewrite, and keeps the insert pending", () => {
    const service = seedWith([rightRule(right.startColumnIndex, 2)]);

    const { result, files } = runLayout({
      tables: { ...insertOnLeft.tables, ...insertOnRight.tables },
    });

    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      stops: [],
      didntLand: [],
      notSent: [
        {
          subject: { table: "layoutRight", path: "newColumns" },
          what: heldBackWhat,
        },
      ],
    });
    expect(result.plan.map(({ subject }) => subject)).toEqual([
      { table: "layoutLeft", path: "newColumns" },
    ]);
    // Left's insert shifted Right one column; Right's own layout is unchanged.
    expect(tabRules(service)).toEqual([
      rightRule(right.startColumnIndex + 1, 2),
    ]);
    const alterations = Val.assert(files, "emitted files").alterations;
    expect(alterations).toContain("Memo");
    expect(alterations).not.toContain("Note");
  });

  it("holds back the right Table's column delete beside a left Table's insert", () => {
    seedWith([rightRule(right.startColumnIndex, 2)]);

    const { result, files } = runLayout({
      tables: { ...insertOnLeft.tables, ...deleteOnRight.tables },
    });

    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      notSent: [
        {
          subject: { table: "layoutRight", column: "entry", path: "deleted" },
          what: heldBackWhat,
        },
      ],
    });
    expect(Val.assert(files, "emitted files").alterations).toContain(
      "layoutRight",
    );
  });

  it("sends the held-back insert and its rule rewrite on the next run, on the shifted origin", () => {
    const service = seedWith([rightRule(right.startColumnIndex, 2)]);
    runLayout({ tables: { ...insertOnLeft.tables, ...insertOnRight.tables } });

    const { result } = runLayout(insertOnRight);

    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      didntLand: [],
      notSent: [],
    });
    expect(tabRules(service)).toEqual([
      rightRule(right.startColumnIndex + 1, 3),
    ]);
  });

  it("sends the held-back column delete and its rule rewrite on the next run, on the shifted origin", () => {
    const service = seedWith([rightRule(right.startColumnIndex, 2)]);
    runLayout({ tables: { ...insertOnLeft.tables, ...deleteOnRight.tables } });

    const { result } = runLayout(deleteOnRight);

    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      didntLand: [],
      notSent: [],
    });
    expect(tabRules(service)).toEqual([
      rightRule(right.startColumnIndex + 1, 1),
    ]);
  });

  it("holds back the right Table's column move beside a left Table's insert", () => {
    seedWith([rightRule(right.startColumnIndex, 2)]);

    const { result, files } = runLayout({
      tables: { ...insertOnLeft.tables, ...moveOnRight.tables },
    });

    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      notSent: [
        {
          subject: { table: "layoutRight", path: "moves[0]" },
          what: heldBackWhat,
        },
      ],
    });
    expect(result.plan.map(({ subject }) => subject)).toEqual([
      { table: "layoutLeft", path: "newColumns" },
    ]);
    expect(Val.assert(files, "emitted files").alterations).toContain("moves");
  });

  it("holds back the right Table's append and its rule's widening beside a left Table's insert", () => {
    const service = seedWith([rightRule(right.startColumnIndex, 2)]);

    const { result, files } = runLayout({
      tables: {
        ...insertOnLeft.tables,
        layoutRight: { newColumns: [{ columns: [{ header: "Memo" }] }] },
      },
    });

    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      notSent: [
        {
          subject: { table: "layoutRight", path: "newColumns" },
          what: heldBackWhat,
        },
      ],
    });
    expect(result.plan.map(({ subject }) => subject)).toEqual([
      { table: "layoutLeft", path: "newColumns" },
    ]);
    expect(tabRules(service)).toEqual([
      rightRule(right.startColumnIndex + 1, 2),
    ]);
    expect(Val.assert(files, "emitted files").alterations).toContain("Memo");
  });

  it("sends a Table's insert and its rule rewrite in the same run when no left neighbour shifts it", () => {
    const service = seedWith([rightRule(right.startColumnIndex, 2)]);

    const { result } = runLayout(insertOnRight);

    expect(result).toMatchObject({ outcome: { kind: "sent" }, notSent: [] });
    expect(tabRules(service)).toEqual([rightRule(right.startColumnIndex, 3)]);
  });

  it("names the held-back insert in a dry run's notices and leaves it off the plan", () => {
    seedWith([rightRule(right.startColumnIndex, 2)]);

    const { result } = SyncCoordinator.init().run({
      isSending: false,
      intendedState: {
        alterations: {
          tables: { ...insertOnLeft.tables, ...insertOnRight.tables },
        },
      },
      previousIdNameMap: previousLayoutMap(),
      forcedTraits: [],
      frame,
    });

    expect(result.notices).toContainEqual({
      kind: "ruleHeldBack",
      subject: { table: "layoutRight", path: "newColumns" },
      shiftedBy: "layoutLeft",
    });
    expect(result.plan.map(({ subject }) => subject)).toEqual([
      { table: "layoutLeft", path: "newColumns" },
    ]);
  });
});
