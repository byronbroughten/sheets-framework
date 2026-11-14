import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { type IdNameMapGeneric } from "../01_SpreadsheetSchema/configReaders/idNameMapTypes";
import { type PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState";
import { expectedSheetLayout } from "../testSupport/expectedSheetLayout";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import { type FakeSheetProperties } from "../testSupport/fakeSheetsService";
import { fakeTableSheet } from "../testSupport/fakeSheetsService/fakeTableSheet";
import { parseIdNameMapFileSource } from "./parseIdNameMapFileSource";
import { SyncCoordinator, type SyncRun } from "./SyncCoordinator";
import {
  columnsByHeader,
  frame,
  itemColumns,
  itemGid,
  notesGid,
  optionalNoteAlteration,
  previousItemMap,
  runItem,
  type RunOptions,
  seedFixture,
} from "./syncCoordinatorTestSupport";

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("SyncCoordinator.run API access", () => {
  function notesSheetAt(startRowIndex?: number): FakeSheetProperties {
    return fakeTableSheet.build({
      sheetId: notesGid,
      title: "Notes",
      tableId: "notes",
      name: "Desk Notes",
      columnConfigs: columnsByHeader(["Note", "Owner"]),
      columnNames: ["Note", "Owner"],
      bodyRows: [{}],
      startRowIndex,
    });
  }

  // The previous map holds Desk Notes with no columns: it has no API access.
  function previousMapWithNotes(): IdNameMapGeneric {
    const map = previousItemMap();
    map.tables.notes = { name: "deskNotes", sheetGid: notesGid };
    map.tabs[String(notesGid)] = {
      name: "notes",
      title: "Notes",
      tableIds: ["notes"],
    };
    return map;
  }

  function runNotes({ isSending, alterations = {} }: RunOptions): SyncRun {
    return SyncCoordinator.init().run({
      isSending,
      intendedState: { alterations },
      previousIdNameMap: previousMapWithNotes(),
      forcedTraits: [],
      frame,
    });
  }

  const grant: PendingAlterations = {
    tables: { deskNotes: { letApiAccess: true } },
  } as PendingAlterations;

  it("grants API access on a send, writing the Table's column IDs and taking it into the configs", () => {
    const { grid } = seedFixture({ extraSheets: [notesSheetAt()] });

    const { result, files } = runNotes({ isSending: true, alterations: grant });

    const emitted = Val.assert(files, "emitted files");
    const columnIds =
      grid.sheet(notesGid).values()[expectedSheetLayout.colIdRowIndex] ?? [];
    expect(columnIds).toHaveLength(2);
    for (const columnId of columnIds) expect(columnId).toMatch(/^c:dnt:/);
    expect(emitted.tableConfigs).toContain('"deskNotes": { "tableId": "notes"');
    expect(
      parseIdNameMapFileSource(emitted.idNameMap).tables.notes?.columns,
    ).toEqual({
      note: columnIds[0],
      owner: columnIds[1],
    });
    expect(emitted.alterations).toContain("makePendingAlterations({})");
    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      landed: [
        {
          kind: "send",
          subject: { table: "deskNotes", path: "letApiAccess" },
          what: "set to true",
        },
      ],
    });
  });

  it("grants API access beside another Table's header rename in one send, taking the granted Table into the configs with its column IDs", () => {
    const { grid } = seedFixture({ extraSheets: [notesSheetAt()] });

    const { result, files } = runNotes({
      isSending: true,
      alterations: {
        tables: {
          deskNotes: { letApiAccess: true },
          item: { columns: { optionalNote: { header: "Remark" } } },
        },
      } as PendingAlterations,
    });

    const emitted = Val.assert(files, "emitted files");
    const columnIds =
      grid.sheet(notesGid).values()[expectedSheetLayout.colIdRowIndex] ?? [];
    expect(columnIds).toHaveLength(2);
    expect(emitted.tableConfigs).toContain('"deskNotes": { "tableId": "notes"');
    const emittedTables = parseIdNameMapFileSource(emitted.idNameMap).tables;
    expect(emittedTables.notes?.columns).toEqual({
      note: columnIds[0],
      owner: columnIds[1],
    });
    expect(emittedTables.item?.columns).toMatchObject({
      remark: itemColumns.optionalNote.columnId,
    });
    expect(result).toMatchObject({ outcome: { kind: "sent" } });
  });

  it("plans a grant on a dry run, writing nothing and leaving the Table out of the configs", () => {
    const { grid, batchUpdateCount } = seedFixture({
      extraSheets: [notesSheetAt()],
    });

    const { result, files } = runNotes({
      isSending: false,
      alterations: grant,
    });

    const emitted = Val.assert(files, "emitted files");
    expect(batchUpdateCount()).toBe(0);
    expect(
      grid.sheet(notesGid).values()[expectedSheetLayout.colIdRowIndex],
    ).toEqual(["", ""]);
    expect(emitted.tableConfigs).not.toContain('"deskNotes"');
    expect(result).toMatchObject({
      outcome: { kind: "ready" },
      plan: [
        {
          subject: { table: "deskNotes", path: "letApiAccess" },
          what: "set to true",
        },
      ],
    });
  });

  it("refuses any other entry on a Table being granted before reading, and lets other Tables' entries share the run", () => {
    const service = seedFixture({ extraSheets: [notesSheetAt()] });

    const { result, files } = runNotes({
      isSending: true,
      alterations: {
        tables: {
          deskNotes: {
            letApiAccess: true,
            columns: { note: { emptyValueAllowed: true } },
          },
          ...optionalNoteAlteration.tables,
        },
      } as PendingAlterations,
    });

    expect(files).toBeUndefined();
    expect(service.getCalls).toEqual([]);
    expect(result).toMatchObject({
      outcome: { kind: "refused" },
      refusals: [
        {
          kind: "besideGrant",
          table: "deskNotes",
          path: "note.emptyValueAllowed",
        },
      ],
    });
  });

  it("refuses an entry on a Table without API access before reading, unless it is a grant", () => {
    const service = seedFixture({ extraSheets: [notesSheetAt()] });

    const { result, files } = runNotes({
      isSending: true,
      alterations: {
        tables: {
          deskNotes: { columns: { note: { emptyValueAllowed: true } } },
        },
      } as PendingAlterations,
    });

    expect(files).toBeUndefined();
    expect(service.getCalls).toEqual([]);
    expect(result).toMatchObject({
      outcome: { kind: "refused" },
      refusals: [
        {
          kind: "noApiAccess",
          table: "deskNotes",
          path: "note.emptyValueAllowed",
        },
      ],
    });
  });

  it("stops a grant on a Table that can't take head rows on row 4, naming it, and writes nothing", () => {
    const { grid, batchUpdateCount } = seedFixture({
      extraSheets: [notesSheetAt(expectedSheetLayout.tableHeaderRowIndex + 2)],
    });
    const before = grid.sheet(notesGid).values();

    const { result, files } = runNotes({ isSending: true, alterations: grant });

    const emitted = Val.assert(files, "emitted files");
    expect(batchUpdateCount()).toBe(0);
    expect(grid.sheet(notesGid).values()).toEqual(before);
    expect(emitted.tableConfigs).not.toContain('"deskNotes"');
    expect(emitted.alterations).toBeUndefined();
    expect(result).toMatchObject({
      outcome: { kind: "stopped" },
      landed: [],
      stops: [
        {
          kind: "headRows",
          table: "deskNotes",
          tab: "Notes",
          headerRow: expectedSheetLayout.tableHeaderRowIndex + 3,
        },
      ],
    });
  });

  it("revokes API access on a send, dropping the Table from the configs without deleting it", () => {
    const { grid, batchUpdateCount } = seedFixture();
    const before = grid.sheet(itemGid).values();

    const { result, files } = runItem({
      isSending: true,
      alterations: { tables: { item: { letApiAccess: false } } },
    });

    const emitted = Val.assert(files, "emitted files");
    expect(batchUpdateCount()).toBe(0);
    expect(grid.sheet(itemGid).values()).toEqual(before);
    expect(emitted.tableConfigs).not.toContain('"item"');
    expect(emitted.columnConfigs).not.toContain('"item"');
    expect(parseIdNameMapFileSource(emitted.idNameMap).tables.item).toEqual({
      name: "item",
      sheetGid: itemGid,
    });
    expect(emitted.alterations).toContain("makePendingAlterations({})");
    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      landed: [
        {
          kind: "send",
          subject: { table: "item", path: "letApiAccess" },
          what: "set to false",
        },
      ],
    });
  });
});
