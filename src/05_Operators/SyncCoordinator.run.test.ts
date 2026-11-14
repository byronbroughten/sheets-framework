import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { type PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState";
import { expectedSheetLayout } from "../testSupport/expectedSheetLayout";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import { parseIdNameMapFileSource } from "./parseIdNameMapFileSource";
import {
  emptyValueAllowedOf,
  itemColumns,
  itemGid,
  optionalNoteAlteration,
  runItem,
  seedFixture,
} from "./syncCoordinatorTestSupport";

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("SyncCoordinator.run", () => {
  it("lands an emptyValueAllowed alteration in the configs, writing nothing to the sheet, and removes it from alterations.ts", () => {
    const { batchUpdateCount } = seedFixture();

    const { result, files } = runItem({
      isSending: true,
      alterations: optionalNoteAlteration,
    });

    const emitted = Val.assert(files, "emitted files");
    expect(itemColumns.optionalNote.emptyValueAllowed).toBe(true);
    expect(emptyValueAllowedOf(emitted, "optionalNote")).toBe(false);
    expect(emitted.alterations).toBe(
      [
        "// sync rewrites this file: each alteration that lands is removed, and comments here won't survive a run.",
        'import { makePendingAlterations } from "@byronbroughten/sheets-framework";',
        "",
        "export const pendingAlterations = makePendingAlterations({});",
        "",
      ].join("\n"),
    );
    expect(batchUpdateCount()).toBe(0);
    expect(result).toMatchObject({
      mode: "send",
      outcome: { kind: "sent", summary: "1 alteration landed" },
      landed: [
        {
          kind: "send",
          subject: {
            table: "item",
            column: "optionalNote",
            path: "emptyValueAllowed",
          },
          what: "set to false",
        },
      ],
    });
  });

  it("refreshes the configs from the live sheet on a dry run, without the planned send, and leaves alterations.ts alone", () => {
    const { batchUpdateCount } = seedFixture({
      item: {
        columns: {
          requiredCount: { ...itemColumns.requiredCount, header: "Tally" },
        },
      },
    });

    const { result, files } = runItem({
      isSending: false,
      alterations: optionalNoteAlteration,
    });

    const emitted = Val.assert(files, "emitted files");
    expect(itemColumns.requiredCount.header).not.toBe("Tally");
    expect(emitted.columnConfigs).toContain('"header": "Tally"');
    expect(emptyValueAllowedOf(emitted, "optionalNote")).toBe(true);
    expect(emitted.alterations).toBeUndefined();
    expect(batchUpdateCount()).toBe(0);
    expect(result).toMatchObject({
      mode: "dry run",
      outcome: { kind: "ready" },
      stops: [],
      plan: [
        {
          subject: {
            table: "item",
            column: "optionalNote",
            path: "emptyValueAllowed",
          },
          what: "set to false",
        },
      ],
    });
  });

  it("writes no column ID on a bare run and leaves an ID-less column out of the configs", () => {
    const { batchUpdateCount, grid } = seedFixture({
      item: { columns: { extra: { columnId: "", header: "Extra" } } },
    });

    const { files } = runItem({ isSending: false });

    const emitted = Val.assert(files, "emitted files");
    expect(batchUpdateCount()).toBe(0);
    expect(
      grid.sheet(itemGid).values()[expectedSheetLayout.colIdRowIndex],
    ).toContain("");
    expect(emitted.columnConfigs).not.toContain('"extra"');
    expect(
      parseIdNameMapFileSource(emitted.idNameMap).tables.item?.columns,
    ).not.toHaveProperty("extra");
  });

  it("sends nothing on a stop, unrelated changes included, yet lists the plan and the conflict", () => {
    const { batchUpdateCount } = seedFixture({
      item: { columns: { optionalNote: { columnId: "", header: "Scratch" } } },
    });

    const { result, files } = runItem({
      isSending: true,
      alterations: {
        tables: {
          item: {
            columns: {
              optionalNote: { emptyValueAllowed: false },
              requiredCount: { emptyValueAllowed: true },
            },
          },
        },
      },
    });

    const emitted = Val.assert(files, "emitted files");
    expect(batchUpdateCount()).toBe(0);
    expect(emitted.alterations).toBeUndefined();
    expect(emptyValueAllowedOf(emitted, "requiredCount")).toBe(false);
    expect(result).toMatchObject({
      mode: "send",
      outcome: { kind: "stopped" },
      landed: [],
      stops: [
        {
          kind: "conflict",
          subject: {
            table: "item",
            column: "optionalNote",
            path: "emptyValueAllowed",
          },
          label: "a person deleted this column",
          isColumnGone: true,
        },
      ],
      plan: [
        {
          subject: {
            table: "item",
            column: "requiredCount",
            path: "emptyValueAllowed",
          },
          what: "set to true",
        },
      ],
    });
  });

  it("refuses an alteration naming a Table the previous map doesn't hold, before reading anything", () => {
    const service = seedFixture();

    const { result, files } = runItem({
      isSending: true,
      alterations: { tables: { noSuchTable: {} } } as PendingAlterations,
    });

    expect(files).toBeUndefined();
    expect(service.getCalls).toEqual([]);
    expect(result).toMatchObject({
      mode: "send",
      outcome: { kind: "refused" },
      refusals: [{ kind: "noSuchTable", entry: "tables", name: "noSuchTable" }],
    });
  });
});
