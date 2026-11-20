import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { type PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState";
import {
  layoutGid,
  layoutOrigins,
  layoutSheet,
} from "../02_SpreadsheetRaw/spreadsheetRawTestSupport";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  type FakeSheetsService,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { SyncCoordinator } from "./SyncCoordinator";
import {
  frame,
  previousLayoutMap,
  runLayout,
} from "./syncCoordinatorTestSupport";

type SheetsRule = GoogleAppsScript.Sheets.Schema.ConditionalFormatRule;

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

function tabRules(service: FakeSheetsService): SheetsRule[] {
  return service.grid.sheet(layoutGid).conditionalFormats ?? [];
}

const amountTrue = {
  columns: ["amount"],
  condition: { type: "CUSTOM_FORMULA" as const, values: ["=TRUE"] },
  format: { backgroundColor: "#ff0000" },
};
const addOnRight: PendingAlterations = {
  tables: { layoutRight: { conditionalFormats: { add: [amountTrue] } } },
};
const insertOnLeft: PendingAlterations = {
  tables: {
    layoutLeft: {
      newColumns: [{ columns: [{ header: "Note" }], before: "amount" }],
    },
  },
};
const deleteOnLeft: PendingAlterations = {
  tables: { layoutLeft: { columns: { amount: { deleted: true } } } },
};
const heldBackLine = {
  subject: { table: "layoutRight", path: "conditionalFormats.add[0]" },
  what: "held back: layoutLeft's column change shifts layoutRight's columns this run, so it is sent next run",
};

function amountColIndexes(service: FakeSheetsService): number[] {
  return tabRules(service).flatMap(({ ranges = [] }) =>
    ranges.map((range) => Val.assert(range.startColumnIndex, "rule column")),
  );
}

describe("SyncCoordinator.run holds back a rule change a left neighbour's column change shifts", () => {
  it("sends a left Table's column insert, holds back the right Table's add and keeps it pending", () => {
    const service = stubSheetsService({ sheets: [layoutSheet()] });

    const { result, files } = runLayout({
      tables: { ...insertOnLeft.tables, ...addOnRight.tables },
    });

    expect(tabRules(service)).toEqual([]);
    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      stops: [],
      notSent: [heldBackLine],
    });
    expect(result.plan.map(({ subject }) => subject)).not.toContainEqual(
      heldBackLine.subject,
    );
    const alterations = Val.assert(files, "emitted files").alterations;
    expect(alterations).toContain("layoutRight");
    expect(alterations).not.toContain("Note");
  });

  it("holds back the right Table's add beside a left Table's column delete", () => {
    const service = stubSheetsService({ sheets: [layoutSheet()] });

    const { result } = runLayout({
      tables: { ...deleteOnLeft.tables, ...addOnRight.tables },
    });

    expect(tabRules(service)).toEqual([]);
    expect(result).toMatchObject({ notSent: [heldBackLine] });
  });

  it("sends the held-back add on the next run, on the shifted layout", () => {
    const service = stubSheetsService({ sheets: [layoutSheet()] });
    runLayout({ tables: { ...deleteOnLeft.tables, ...addOnRight.tables } });

    const { result } = runLayout(addOnRight);

    expect(result).toMatchObject({ outcome: { kind: "sent" }, notSent: [] });
    // Left lost a column, so Right's amount moved one column left.
    expect(amountColIndexes(service)).toEqual([
      layoutOrigins.layoutRight.startColumnIndex,
    ]);
  });

  // The header zone puts every managed Table's header on one row, so a left neighbour always shares rows; one to the right shifts nothing.
  it("sends the add when the Table changing its columns sits to its right", () => {
    const service = stubSheetsService({ sheets: [layoutSheet()] });

    const { result } = runLayout({
      tables: {
        layoutRight: {
          newColumns: [{ columns: [{ header: "Note" }], before: "amount" }],
        },
        layoutLeft: { conditionalFormats: { add: [amountTrue] } },
      },
    });

    expect(result).toMatchObject({ outcome: { kind: "sent" }, notSent: [] });
    expect(amountColIndexes(service)).toEqual([
      layoutOrigins.layoutLeft.startColumnIndex + 1,
    ]);
  });

  it("names the held-back add in a dry run's notices", () => {
    stubSheetsService({ sheets: [layoutSheet()] });

    const { result } = SyncCoordinator.init().run({
      isSending: false,
      intendedState: {
        alterations: {
          tables: { ...insertOnLeft.tables, ...addOnRight.tables },
        },
      },
      previousIdNameMap: previousLayoutMap(),
      forcedTraits: [],
      frame,
    });

    expect(result.notices).toContainEqual({
      kind: "ruleHeldBack",
      subject: heldBackLine.subject,
      shiftedBy: "layoutLeft",
    });
    expect(result.plan.map(({ subject }) => subject)).not.toContainEqual(
      heldBackLine.subject,
    );
  });
});
