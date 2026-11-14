import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { type PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState";
import { expectedSheetLayout } from "../testSupport/expectedSheetLayout";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  type FakeCellValue,
  type FakeSheetsService,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { parseIdNameMapFileSource } from "./parseIdNameMapFileSource";
import {
  bodyRowsWithNote,
  itemColumns,
  itemGid,
  itemSheet,
  runItem,
  seedFixture,
} from "./syncCoordinatorTestSupport";

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("SyncCoordinator.run column deletion", () => {
  const deleteOptionalNote: PendingAlterations = {
    tables: { item: { columns: { optionalNote: { deleted: true } } } },
  };
  const noteSubject = {
    table: "item",
    column: "optionalNote",
    path: "deleted",
  };
  const noteHeader = itemColumns.optionalNote.header;

  function itemSheetRows(service: FakeSheetsService): FakeCellValue[][] {
    return service.grid.sheet(itemGid).values();
  }

  it("plans a deletion on a dry run, sending nothing", () => {
    const service = seedFixture();

    const { result } = runItem({
      isSending: false,
      alterations: deleteOptionalNote,
    });

    expect(service.batchUpdateCount()).toBe(0);
    expect(result).toMatchObject({
      mode: "dry run",
      outcome: { kind: "ready" },
      refusals: [],
      plan: [{ subject: noteSubject, what: "delete the column" }],
    });
  });

  it("deletes the column on a send, from its column ID down, and drops it from the configs, the ID-name map and alterations.ts", () => {
    const service = seedFixture({ item: { bodyRows: bodyRowsWithNote } });
    const before = itemSheetRows(service);
    const { colIdRowIndex, tableHeaderRowIndex } = expectedSheetLayout;

    const { result, files } = runItem({
      isSending: true,
      alterations: deleteOptionalNote,
    });

    const emitted = Val.assert(files, "emitted files");
    const after = itemSheetRows(service);
    expect(after[tableHeaderRowIndex]).not.toContain(noteHeader);
    expect(after[colIdRowIndex]).not.toContain(
      itemColumns.optionalNote.columnId,
    );
    expect(after.flat()).not.toContain("kept");
    expect(after[tableHeaderRowIndex]?.filter((cell) => cell !== "")).toEqual(
      before[tableHeaderRowIndex]?.filter(
        (cell) => cell !== "" && cell !== noteHeader,
      ),
    );
    expect(emitted.columnConfigs).not.toContain('"optionalNote"');
    expect(
      parseIdNameMapFileSource(emitted.idNameMap).tables.item?.columns,
    ).not.toHaveProperty("optionalNote");
    expect(emitted.alterations).toContain("makePendingAlterations({})");
    expect(result).toMatchObject({
      mode: "send",
      outcome: { kind: "sent" },
      landed: [{ kind: "send", subject: noteSubject }],
      didntLand: [],
    });
  });

  it("sends the deletion in a flush after the run's other writes on its Table", () => {
    const service = stubSheetsService({
      sheets: [
        itemSheet({ columns: { extra: { columnId: "", header: "Extra" } } }),
      ],
    });

    const { result, files } = runItem({
      isSending: true,
      alterations: deleteOptionalNote,
    });

    const emitted = Val.assert(files, "emitted files");
    const headerRow =
      itemSheetRows(service)[expectedSheetLayout.tableHeaderRowIndex];
    expect(headerRow).toContain("Extra");
    expect(headerRow).not.toContain(noteHeader);
    expect(emitted.columnConfigs).toContain('"extra"');
    expect(emitted.columnConfigs).not.toContain('"optionalNote"');
    expect(result.outcome).toMatchObject({ kind: "sent" });
  });
});
