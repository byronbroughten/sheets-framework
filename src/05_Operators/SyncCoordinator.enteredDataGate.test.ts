import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { HeldFlushRefusal } from "../00_Source/RawSource/HeldFlushRefusal";
import { type PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState";
import { expectedSheetLayout } from "../testSupport/expectedSheetLayout";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  type FakeSheetsService,
  type FakeSheetsServiceOptions,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { SyncCoordinator, type SyncRun } from "./SyncCoordinator";
import {
  frame,
  itemAtEachFlush,
  type ItemAtFlush,
  type ItemAtFlushProps,
  itemGid,
  itemSheet,
  noteColIndex,
  optionalNoteAlteration,
  previousItemMap,
  runItem,
} from "./syncCoordinatorTestSupport";

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("SyncCoordinator.run entered-data gate", () => {
  const extraColumn = { extra: { columnId: "", header: "Extra" } };

  function seedWithExtra(
    options: Omit<FakeSheetsServiceOptions, "sheets"> = {},
  ): FakeSheetsService {
    return stubSheetsService({
      sheets: [itemSheet({ columns: extraColumn })],
      ...options,
    });
  }

  function dryRun(isHostWithholding: boolean): SyncRun {
    return SyncCoordinator.init().run({
      isSending: false,
      intendedState: { alterations: {} },
      previousIdNameMap: previousItemMap(),
      forcedTraits: [],
      frame,
      isHostWithholding,
    });
  }

  function refusingSend(refusal: Error): SyncRun {
    seedWithExtra({
      beforeEachBatchUpdate: () => {
        throw refusal;
      },
    });
    return runItem({ isSending: true });
  }

  it("flushes an unstopped dry run's apply after its files are built when the host withholds", () => {
    const { batchUpdateCount, grid } = seedWithExtra({ isDryRun: true });

    const { result, files } = dryRun(true);

    const emitted = Val.assert(files, "emitted files");
    expect(batchUpdateCount()).toBe(1);
    expect(emitted.columnConfigs).not.toContain('"extra"');
    expect(
      grid.sheet(itemGid).values()[expectedSheetLayout.colIdRowIndex],
    ).toContain("");
    expect(result).toMatchObject({
      mode: "dry run",
      outcome: { kind: "ready" },
    });
  });

  it("sends nothing on a dry run without the withholding flag", () => {
    const { batchUpdateCount } = seedWithExtra();

    dryRun(false);

    expect(batchUpdateCount()).toBe(0);
  });

  it("flushes nothing on a stopped dry run, even when the host withholds", () => {
    const { batchUpdateCount } = stubSheetsService({
      sheets: [
        itemSheet({
          columns: { optionalNote: { columnId: "", header: "Scratch" } },
        }),
      ],
      isDryRun: true,
    });

    const { result } = SyncCoordinator.init().run({
      isSending: false,
      intendedState: { alterations: optionalNoteAlteration },
      previousIdNameMap: previousItemMap(),
      forcedTraits: [],
      frame,
      isHostWithholding: true,
    });

    expect(result.outcome).toEqual({ kind: "stopped" });
    expect(batchUpdateCount()).toBe(0);
  });

  it("stops a send the gate refuses before any flush landed, writing no files", () => {
    const { result, files } = refusingSend(
      new HeldFlushRefusal({
        kind: "uncovered",
        landedFlushes: 0,
        flushCount: undefined,
      }),
    );

    expect(files).toBeUndefined();
    expect(result).toMatchObject({
      mode: "send",
      outcome: { kind: "stoppedBeforeSending", heldBy: "uncovered" },
      landed: [],
    });
  });

  it("stops a send partway when the gate refuses after a landed flush, writing no files", () => {
    const { result, files } = refusingSend(
      new HeldFlushRefusal({
        kind: "mismatched",
        landedFlushes: 1,
        flushCount: undefined,
      }),
    );

    expect(files).toBeUndefined();
    expect(result.outcome).toEqual({
      kind: "stoppedPartway",
      landedFlushes: 1,
      flushCount: undefined,
      heldBy: "mismatched",
    });
  });

  describe("a deleted column's cells", () => {
    const deleteOptionalNote: PendingAlterations = {
      tables: { item: { columns: { optionalNote: { deleted: true } } } },
    };
    function noteAtFirstFlush(
      props: Pick<ItemAtFlushProps, "isSending" | "isHostWithholding">,
    ): ItemAtFlush | undefined {
      return itemAtEachFlush({
        ...props,
        alterations: deleteOptionalNote,
        columns: extraColumn,
      })[0];
    }

    it("reads the column before a send's flush, holding the cell a person entered", () => {
      expect(noteAtFirstFlush({ isSending: true })).toEqual({
        colIndexesRead: expect.arrayContaining([noteColIndex]),
        note: "kept",
      });
    });

    it("reads the column before a withheld dry run's flush", () => {
      expect(
        noteAtFirstFlush({ isSending: false, isHostWithholding: true }),
      ).toEqual({
        colIndexesRead: expect.arrayContaining([noteColIndex]),
        note: "kept",
      });
    });
  });

  it("passes any other flush error on up", () => {
    expect(() => refusingSend(new Error("Sheets API 500"))).toThrow(
      "Sheets API 500",
    );
  });
});
