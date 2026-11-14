import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { HeldFlushRefusal } from "../00_Source/RawSource/HeldFlushRefusal";
import { installedConfigs } from "../01_SpreadsheetSchema/configReaders/configRegister";
import type { IdNameMapGeneric } from "../01_SpreadsheetSchema/configReaders/idNameMapTypes";
import type { PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState";
import type {
  ForcedTrait,
  SyncRunFrame,
} from "../01_SpreadsheetSchema/configReaders/syncRunResult";
import { getTableTraitByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { IdPrefixesByTableId } from "../01_SpreadsheetSchema/idPrefixes";
import { expectedSheetLayout } from "../testSupport/expectedSheetLayout";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  buildGridRows,
  type FakeCell,
  type FakeCellValue,
  type FakeSheetProperties,
  type FakeSheetsService,
  type FakeSheetsServiceOptions,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import {
  type FakeBodyRow,
  type FakeTableColumn,
  fakeTableSheet,
} from "../testSupport/fakeSheetsService/fakeTableSheet";
import { parseIdNameMapFileSource } from "./parseIdNameMapFileSource";
import {
  type ConfigRegeneration,
  SyncCoordinator,
  type SyncRun,
} from "./SyncCoordinator";
import { TableConfigsOperator } from "./TableConfigsOperator";

type GridRange = GoogleAppsScript.Sheets.Schema.GridRange;

const { columnConfigs } = installedConfigs();
const itemGid = getTableTraitByName("item", "sheetGid");
const itemColumns = columnConfigs.item;
type ItemColumnName = keyof typeof itemColumns;

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

interface ItemSheetOptions {
  columns?: Record<string, FakeTableColumn>; // Replaces a configured column by name, or adds one after them.
  startRowIndex?: number;
  columnTypes?: Record<number, string>; // By absolute column index.
  name?: string; // The Table's Sheets name; "Item" when unset.
  columnNames?: string[]; // The sheet's column order; the configured order when unset.
  bodyRows?: FakeBodyRow<string>[]; // One row with an ID when unset.
  actionRowCells?: FakeBodyRow<string>; // A Fill Row IDs checkbox above ID when unset.
}

const itemRowId = "r:itm:fixture";

// The dev Item Table as the previous configs record it, with its ID filled and its checkbox in place, so every run starts quiet.
function itemSheet(options: ItemSheetOptions = {}): FakeSheetProperties {
  const sheetColumns: Record<string, FakeTableColumn> = {
    ...itemColumns,
    ...options.columns,
  };
  return fakeTableSheet.build({
    sheetId: itemGid,
    title: "Item",
    tableId: "item",
    name: options.name ?? "Item",
    columnConfigs: sheetColumns,
    columnNames: options.columnNames ?? Object.keys(sheetColumns),
    bodyRows: options.bodyRows ?? [{ id: itemRowId }],
    actionRowCells: options.actionRowCells ?? { id: false },
    startRowIndex: options.startRowIndex,
    columnTypes: options.columnTypes,
  });
}

function seedFixture(
  options: {
    item?: ItemSheetOptions;
    extraSheets?: FakeSheetProperties[];
  } = {},
): FakeSheetsService {
  return stubSheetsService({
    sheets: [itemSheet(options.item), ...(options.extraSheets ?? [])],
  });
}

// Keyed by header, so a fixture names each column by what its sheet shows.
function columnsByHeader(
  headers: readonly string[],
  columnIds: readonly string[] = [],
): Record<string, FakeTableColumn> {
  return Object.fromEntries(
    headers.map((header, colIndex) => [
      header,
      { columnId: columnIds[colIndex] ?? "", header },
    ]),
  );
}

// A Table the previous configs don't hold, at a GID they don't hold either.
const notesGid = 999001;

function notesSheet(): FakeSheetProperties {
  return fakeTableSheet.build({
    sheetId: notesGid,
    title: "Notes",
    tableId: "notes",
    name: "Desk Notes",
    columnConfigs: columnsByHeader(["Note"]),
    columnNames: ["Note"],
    bodyRows: [{}],
  });
}

interface TableConfigsEntryFixture {
  tableKey: string;
  tableId: string;
  tableName: string;
  sheetGid: number;
  idPrefix: string;
  hasIdColumn: boolean;
}

// Every fixture Table sits at the expected origin and carries a "Name" header.
function tableConfigsEntry({
  tableKey,
  tableId,
  tableName,
  sheetGid,
  idPrefix,
  hasIdColumn,
}: TableConfigsEntryFixture): string {
  return `"${tableKey}": { "tableId": "${tableId}", "tableName": "${tableName}", "sheetGid": ${sheetGid}, "idPrefix": "${idPrefix}", "hasIdColumn": ${hasIdColumn}, "hasNameColumn": true }`;
}

const frame: SyncRunFrame = {
  target: "dev",
  command: "npm run dev:sync",
  generatedDir: "dev/generated",
  alterationsPath: "dev/alterations.ts",
};

// The dev Item Table as the previous ID-name map records it.
function previousItemMap(): IdNameMapGeneric {
  return {
    tables: {
      item: {
        name: "item",
        sheetGid: itemGid,
        columns: Object.fromEntries(
          Object.entries(itemColumns).map(([columnName, { columnId }]) => [
            columnName,
            columnId,
          ]),
        ),
      },
    },
    tabs: {
      [String(itemGid)]: { name: "item", title: "Item", tableIds: ["item"] },
    },
  };
}

interface RunOptions {
  isSending: boolean;
  alterations?: PendingAlterations;
  forcedTraits?: ForcedTrait[];
}

function runItem({
  isSending,
  alterations = {},
  forcedTraits = [],
}: RunOptions): SyncRun {
  return SyncCoordinator.init().run({
    isSending,
    intendedState: { alterations },
    previousIdNameMap: previousItemMap(),
    forcedTraits,
    frame,
  });
}

// A send with empty intended state, so each run is the regen alone.
function runWithNoAlterations(coordinator = SyncCoordinator.init()): SyncRun {
  return coordinator.run({
    isSending: true,
    intendedState: { alterations: {} },
    previousIdNameMap: { tables: {}, tabs: {} },
    forcedTraits: [],
    frame,
  });
}

function sendWithNoAlterations(
  coordinator = SyncCoordinator.init(),
): ConfigRegeneration {
  return Val.assert(
    runWithNoAlterations(coordinator).files,
    "files of a send with no alterations",
  );
}

// The emitted line of one Item column, so a test reads the trait the configs carry.
function emittedItemColumn(
  files: ConfigRegeneration,
  columnName: string,
): string {
  const line = files.columnConfigs
    .split("\n")
    .find((sourceLine) => sourceLine.startsWith(`    "${columnName}": `));
  return Val.assert(line, `emitted ${columnName} column`);
}

function emptyValueAllowedOf(
  files: ConfigRegeneration,
  columnName: string,
): unknown {
  const match = emittedItemColumn(files, columnName).match(
    /"emptyValueAllowed": (true|false)/,
  );
  return JSON.parse(Val.assert(match?.[1], "emitted emptyValueAllowed"));
}

const optionalNoteAlteration: PendingAlterations = {
  tables: {
    item: { columns: { optionalNote: { emptyValueAllowed: false } } },
  },
};

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
    const noteColIndex = Object.keys(itemColumns).indexOf("optionalNote");

    interface NoteAtFlush {
      isNoteColumnRead: boolean;
      note: unknown;
    }

    // One column's own range, from any row; the head-row reads span every column.
    function readsNoteColumn(service: FakeSheetsService): boolean {
      return service.getByDataFilterCalls.some((resource) =>
        (
          (resource as { dataFilters?: { gridRange?: GridRange }[] })
            .dataFilters ?? []
        ).some(
          ({ gridRange }) =>
            gridRange?.sheetId === itemGid &&
            gridRange.startColumnIndex === noteColIndex &&
            gridRange.endColumnIndex === noteColIndex + 1,
        ),
      );
    }

    // The note sits below the top row, which sync's head-row read already covers.
    function heldNoteBeforeFlush(props: {
      isSending: boolean;
      isHostWithholding?: boolean;
    }): NoteAtFlush[] {
      const running: { coordinator?: SyncCoordinator } = {};
      const heldAtFlush: NoteAtFlush[] = [];
      const service = stubSheetsService({
        sheets: [
          itemSheet({
            columns: extraColumn,
            bodyRows: [
              { id: itemRowId },
              { id: "r:itm:second", optionalNote: "kept" },
            ],
          }),
        ],
        isDryRun: !props.isSending,
        beforeEachBatchUpdate: () => {
          heldAtFlush.push({
            isNoteColumnRead: readsNoteColumn(service),
            note: Val.assert(running.coordinator, "the running coordinator")
              .ss.raw.table("item")
              .row(1)
              .cell(noteColIndex)
              .valueOrEmpty(),
          });
        },
      });
      running.coordinator = SyncCoordinator.init();
      running.coordinator.run({
        ...props,
        intendedState: { alterations: deleteOptionalNote },
        previousIdNameMap: previousItemMap(),
        forcedTraits: [],
        frame,
      });
      return heldAtFlush;
    }

    it("reads the column before a send's flush, holding the cell a person entered", () => {
      expect(heldNoteBeforeFlush({ isSending: true })[0]).toEqual({
        isNoteColumnRead: true,
        note: "kept",
      });
    });

    it("reads the column before a withheld dry run's flush", () => {
      expect(
        heldNoteBeforeFlush({ isSending: false, isHostWithholding: true })[0],
      ).toEqual({ isNoteColumnRead: true, note: "kept" });
    });
  });

  it("passes any other flush error on up", () => {
    expect(() => refusingSend(new Error("Sheets API 500"))).toThrow(
      "Sheets API 500",
    );
  });
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
    const service = seedFixture({
      item: {
        bodyRows: [
          { id: itemRowId },
          { id: "r:itm:second", optionalNote: "kept" },
        ],
      },
    });
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

describe("SyncCoordinator.run new columns", () => {
  const newColumnsSubject = { table: "item", path: "newColumns" };
  const mintedId = /^c:itm:[0-9A-Za-z_-]{7}$/;

  function itemSheetRows(service: FakeSheetsService): FakeCellValue[][] {
    return service.grid.sheet(itemGid).values();
  }

  function headersOf(rows: FakeCellValue[][]): FakeCellValue[] {
    const headerRow = rows[expectedSheetLayout.tableHeaderRowIndex] ?? [];
    return headerRow.filter((cell) => cell !== "");
  }

  function columnIdUnder(
    rows: FakeCellValue[][],
    header: string,
  ): FakeCellValue | undefined {
    const { colIdRowIndex, tableHeaderRowIndex } = expectedSheetLayout;
    const colIndex = rows[tableHeaderRowIndex]?.indexOf(header) ?? -1;
    return rows[colIdRowIndex]?.[colIndex];
  }

  function itemTableEnd(service: FakeSheetsService): number {
    const range = service.grid.sheet(itemGid).tables[0]?.range;
    return Val.assert(range?.endColumnIndex, "Item Table end");
  }

  it("plans each entry on a dry run, sending nothing", () => {
    const service = seedFixture();
    const [, secondColumn] = Object.keys(itemColumns) as ItemColumnName[];

    const { result } = runItem({
      isSending: false,
      alterations: {
        tables: {
          item: {
            newColumns: [
              { header: "Late fee" },
              { header: "Ref", before: Val.assert(secondColumn, "column") },
            ],
          },
        },
      },
    });

    expect(service.batchUpdateCount()).toBe(0);
    expect(result).toMatchObject({
      outcome: { kind: "ready" },
      plan: [
        {
          subject: newColumnsSubject,
          what: 'create "Late fee" at the Table\'s end',
        },
        {
          subject: newColumnsSubject,
          what: `create "Ref" before ${secondColumn}`,
        },
      ],
    });
  });

  it("creates a column at the Table's end, widening the Table, and the configs and ID-name map take its minted ID", () => {
    const service = seedFixture({
      item: { bodyRows: [{ id: itemRowId }, { id: "r:itm:second" }] },
    });
    const before = itemSheetRows(service);
    const endBefore = itemTableEnd(service);

    const { result, files } = runItem({
      isSending: true,
      alterations: {
        tables: { item: { newColumns: [{ header: "Late fee" }] } },
      },
    });

    const emitted = Val.assert(files, "emitted files");
    const after = itemSheetRows(service);
    expect(headersOf(after)).toEqual([...headersOf(before), "Late fee"]);
    expect(itemTableEnd(service)).toBe(endBefore + 1);
    const columnId = String(columnIdUnder(after, "Late fee"));
    expect(columnId).toMatch(mintedId);
    expect(emittedItemColumn(emitted, "lateFee")).toContain(
      `"columnId": "${columnId}"`,
    );
    expect(emptyValueAllowedOf(emitted, "lateFee")).toBe(false);
    expect(
      parseIdNameMapFileSource(emitted.idNameMap).tables.item?.columns,
    ).toMatchObject({ lateFee: columnId });
    expect(emitted.alterations).toContain("makePendingAlterations({})");
    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      landed: [
        {
          kind: "createdColumn",
          table: "item",
          header: "Late fee",
          column: "lateFee",
          columnId,
          position: "at the Table's end",
        },
      ],
    });
  });

  it("creates a column before the one an entry names, inside the Table, keeping the empty value allowed it sets", () => {
    const service = seedFixture();
    const [, secondColumn] = Object.keys(itemColumns) as ItemColumnName[];
    const secondName = Val.assert(secondColumn, "second column");
    const before = itemSheetRows(service);
    const endBefore = itemTableEnd(service);

    const { files } = runItem({
      isSending: true,
      alterations: {
        tables: {
          item: {
            newColumns: [
              { header: "Ref", before: secondName, emptyValueAllowed: true },
            ],
          },
        },
      },
    });

    const emitted = Val.assert(files, "emitted files");
    const after = itemSheetRows(service);
    const [firstHeader, ...restHeaders] = headersOf(before);
    expect(headersOf(after)).toEqual([firstHeader, "Ref", ...restHeaders]);
    expect(itemTableEnd(service)).toBe(endBefore + 1);
    const columnId = String(columnIdUnder(after, "Ref"));
    expect(columnId).toMatch(mintedId);
    expect(emittedItemColumn(emitted, "ref")).toContain(
      `"columnId": "${columnId}"`,
    );
    expect(emptyValueAllowedOf(emitted, "ref")).toBe(true);
  });

  it("throws before sending anything on an entry placed before the Table's first column", () => {
    const service = seedFixture();
    const [firstColumn] = Object.keys(itemColumns) as ItemColumnName[];
    const firstName = Val.assert(firstColumn, "first column");

    expect(() =>
      runItem({
        isSending: true,
        alterations: {
          tables: {
            item: {
              newColumns: [
                { header: "Late fee" },
                { header: "Ref", before: firstName },
              ],
            },
          },
        },
      }),
    ).toThrow(/its first column/);
    expect(service.batchUpdateCount()).toBe(0);
  });

  it("stops before sending on an entry whose header the Table already has, naming the column and its ID", () => {
    const service = seedFixture();
    const { header, columnId } = itemColumns.optionalNote;

    const { result, files } = runItem({
      isSending: true,
      alterations: {
        tables: {
          item: { newColumns: [{ header }, { header: "Late fee" }] },
        },
      },
    });

    expect(service.batchUpdateCount()).toBe(0);
    expect(files?.alterations).toBeUndefined();
    expect(result).toMatchObject({
      outcome: { kind: "stopped" },
      stops: [
        {
          kind: "headerTaken",
          table: "item",
          header,
          column: "optionalNote",
          columnId,
        },
      ],
    });
  });
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

describe("SyncCoordinator.run typed trait changes", () => {
  // Item's requiredCount is the second column, recorded as DOUBLE.
  const itemTypes = { 0: "TEXT", 1: "DOUBLE", 2: "TEXT", 3: "TEXT" };

  it("lists nothing on a quiet run", () => {
    seedFixture({ item: { columnTypes: itemTypes } });

    const { result } = runItem({ isSending: false });

    expect(result.configs.typedTraitChanges).toEqual([]);
  });

  it("lists a person's column type change with from and to", () => {
    seedFixture({ item: { columnTypes: { ...itemTypes, 1: "CURRENCY" } } });

    const { result } = runItem({ isSending: false });

    expect(result.configs.typedTraitChanges).toEqual([
      {
        table: "item",
        column: "requiredCount",
        trait: "columnType",
        from: "DOUBLE",
        to: "CURRENCY",
      },
    ]);
  });

  it("doesn't list a sent alteration", () => {
    seedFixture({ item: { columnTypes: itemTypes } });

    const { result } = runItem({
      isSending: true,
      alterations: optionalNoteAlteration,
    });

    expect(result.configs.typedTraitChanges).toEqual([]);
  });
});

describe("SyncCoordinator.run force", () => {
  const rename: PendingAlterations = {
    tables: { item: { columns: { optionalNote: { header: "Mine" } } } },
  };
  const forcedTraits = [
    { table: "item", column: "optionalNote", trait: "header" },
  ];
  // A person renamed the column on the sheet after the last sync.
  const personRenamed = {
    item: {
      columns: {
        optionalNote: { ...itemColumns.optionalNote, header: "Theirs" },
      },
    },
  };

  it("stops on a trait both sides changed, without the flag", () => {
    seedFixture(personRenamed);

    const { result } = runItem({ isSending: true, alterations: rename });

    expect(result).toMatchObject({
      outcome: { kind: "stopped" },
      forcedTraits: [],
    });
  });

  it("sends a forced trait over the sheet's value and holds it, for that run only", () => {
    seedFixture(personRenamed);

    const forced = runItem({
      isSending: true,
      alterations: rename,
      forcedTraits,
    });

    expect(forced.result).toMatchObject({
      outcome: { kind: "sent", summary: "1 alteration landed" },
      stops: [],
      forcedTraits,
      landed: [
        {
          subject: { table: "item", column: "optionalNote", path: "header" },
          what: "set to Mine, over the sheet's Theirs",
        },
      ],
    });
    expect(forced.files?.alterations).toBeDefined();

    seedFixture(personRenamed);
    const next = runItem({ isSending: true, alterations: rename });
    expect(next.result).toMatchObject({ outcome: { kind: "stopped" } });
  });
});

describe("SyncCoordinator.run renames", () => {
  // A typed column on every Item column, so a test sees the types survive the rename.
  const itemTypes = { 0: "TEXT", 1: "DOUBLE", 2: "TEXT", 3: "TEXT" };
  const optionalNoteIndex = Object.keys(itemColumns).indexOf("optionalNote");
  const emptied = [
    "// sync rewrites this file: each alteration that lands is removed, and comments here won't survive a run.",
    'import { makePendingAlterations } from "@byronbroughten/sheets-framework";',
    "",
    "export const pendingAlterations = makePendingAlterations({});",
    "",
  ].join("\n");

  function liveItemTable(
    service: FakeSheetsService,
  ): GoogleAppsScript.Sheets.Schema.Table {
    return Val.assert(service.grid.sheet(itemGid).tables[0], "live Item Table");
  }

  it("renames a column by writing its header cell, keeping its column type, and the configs follow the new header", () => {
    const service = seedFixture({ item: { columnTypes: itemTypes } });

    const { result, files } = runItem({
      isSending: true,
      alterations: {
        tables: { item: { columns: { optionalNote: { header: "Remark" } } } },
      },
    });

    const emitted = Val.assert(files, "emitted files");
    expect(
      liveItemTable(service).columnProperties?.[optionalNoteIndex],
    ).toMatchObject({ columnName: "Remark", columnType: "TEXT" });
    expect(liveItemTable(service).columnProperties?.[1]).toMatchObject({
      columnType: "DOUBLE",
    });
    expect(emittedItemColumn(emitted, "remark")).toContain(
      `"columnId": "${itemColumns.optionalNote.columnId}", "header": "Remark"`,
    );
    expect(emitted.columnConfigs).not.toContain('"optionalNote"');
    expect(
      parseIdNameMapFileSource(emitted.idNameMap).tables.item?.columns,
    ).toMatchObject({ remark: itemColumns.optionalNote.columnId });
    expect(emitted.alterations).toBe(emptied);
    expect(result).toMatchObject({
      mode: "send",
      outcome: { kind: "sent", summary: "1 alteration landed" },
      landed: [
        {
          kind: "send",
          subject: { table: "item", column: "optionalNote", path: "header" },
          what: "set to Remark",
        },
      ],
      didntLand: [],
    });
  });

  it("renames a Table, keeping its column types, and the configs key it by its new name", () => {
    // The name the previous configs record, so the rename isn't a person's change.
    const service = seedFixture({
      item: {
        columnTypes: itemTypes,
        name: getTableTraitByName("item", "tableName"),
      },
    });

    const { result, files } = runItem({
      isSending: true,
      alterations: { tables: { item: { name: "Stock" } } },
    });

    const emitted = Val.assert(files, "emitted files");
    expect(liveItemTable(service)).toMatchObject({ name: "Stock" });
    expect(liveItemTable(service).columnProperties?.[1]).toMatchObject({
      columnType: "DOUBLE",
    });
    expect(emitted.tableConfigs).toContain(
      '"stock": { "tableId": "item", "tableName": "Stock"',
    );
    expect(
      parseIdNameMapFileSource(emitted.idNameMap).tables.item,
    ).toMatchObject({ name: "stock" });
    expect(emitted.alterations).toBe(emptied);
    expect(result).toMatchObject({
      outcome: { kind: "sent", summary: "1 alteration landed" },
      landed: [
        {
          kind: "send",
          subject: { table: "item", path: "name" },
          what: "set to Stock",
        },
      ],
      didntLand: [],
    });
  });

  it("reports a trait Sheets kept another value for under Didn't land, the configs taking that value and the entry staying pending", () => {
    const service = stubSheetsService({
      sheets: [
        itemSheet({
          columnTypes: itemTypes,
          name: getTableTraitByName("item", "tableName"),
        }),
      ],
      afterEachBatchUpdate(spreadsheet) {
        const table = Val.assert(spreadsheet.sheets[0]?.tables[0], "Item");
        table.name = "Item";
      },
    });

    const { result, files } = runItem({
      isSending: true,
      alterations: {
        tables: {
          item: {
            name: "Stock",
            columns: { optionalNote: { header: "Remark" } },
          },
        },
      },
    });

    const emitted = Val.assert(files, "emitted files");
    expect(liveItemTable(service)).toMatchObject({ name: "Item" });
    expect(
      liveItemTable(service).columnProperties?.[optionalNoteIndex],
    ).toMatchObject({ columnName: "Remark" });
    expect(emitted.tableConfigs).toContain(
      '"item": { "tableId": "item", "tableName": "Item"',
    );
    expect(emittedItemColumn(emitted, "remark")).toContain(
      '"header": "Remark"',
    );
    expect(emitted.alterations).toBe(
      [
        "// sync rewrites this file: each alteration that lands is removed, and comments here won't survive a run.",
        'import { makePendingAlterations } from "@byronbroughten/sheets-framework";',
        "",
        "export const pendingAlterations = makePendingAlterations({",
        '  "tables": {',
        '    "item": { "name": "Stock" }',
        "  }",
        "});",
        "",
      ].join("\n"),
    );
    expect(result).toMatchObject({
      mode: "send",
      outcome: {
        kind: "sent",
        summary: "1 alteration landed; 1 didn't land",
      },
      landed: [
        {
          kind: "send",
          subject: { table: "item", column: "optionalNote", path: "header" },
          what: "set to Remark",
        },
      ],
      didntLand: [
        {
          subject: { table: "item", path: "name" },
          what: "set to Stock; Sheets kept Item",
        },
      ],
    });
  });

  it("sends neither rename on a dry run", () => {
    const service = seedFixture({ item: { columnTypes: itemTypes } });

    runItem({
      isSending: false,
      alterations: {
        tables: {
          item: {
            name: "Stock",
            columns: { optionalNote: { header: "Remark" } },
          },
        },
      },
    });

    expect(service.batchUpdateCount()).toBe(0);
    expect(liveItemTable(service)).toMatchObject({ name: "Item" });
    expect(
      liveItemTable(service).columnProperties?.[optionalNoteIndex],
    ).toMatchObject({ columnName: itemColumns.optionalNote.header });
  });
});

describe("SyncCoordinator.run formulas", () => {
  const optionalNoteIndex = Object.keys(itemColumns).indexOf("optionalNote");
  const formulaAlteration: PendingAlterations = {
    tables: { item: { columns: { optionalNote: { formula: "=${@id}" } } } },
  };

  it("fills every body row with the compiled formula, and the configs record it in brace form", () => {
    const service = seedFixture({
      item: {
        // The name the previous configs record, which the compiled reference uses.
        name: getTableTraitByName("item", "tableName"),
        bodyRows: [{ id: "r:itm:one" }, { id: "r:itm:two" }],
      },
    });

    const { files } = runItem({
      isSending: true,
      alterations: formulaAlteration,
    });

    const noteCells = service.grid
      .sheet(itemGid)
      .bodyRows()
      .map((row) => row[optionalNoteIndex]);
    const compiled = {
      isFormula: true,
      value: `=SINGLE(${getTableTraitByName("item", "tableName")}[${itemColumns.id.header}])`,
    };
    expect(noteCells).toEqual([compiled, compiled]);
    expect(
      emittedItemColumn(Val.assert(files, "emitted files"), "optionalNote"),
    ).toContain('"formula": "=${@id}"');
  });

  it("refuses a brace naming no column before any read", () => {
    const service = seedFixture();

    const { result } = runItem({
      isSending: true,
      alterations: {
        tables: {
          item: { columns: { optionalNote: { formula: "=${@nope}" } } },
        },
      },
    });

    expect(result.refusals).toEqual([
      expect.objectContaining({ kind: "unknownBrace", brace: "${@nope}" }),
    ]);
    expect(service.getCalls).toEqual([]);
  });
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

describe("SyncCoordinator.run Fill Row IDs", () => {
  const itemColumnNames = Object.keys(itemColumns);
  const idColIndex = itemColumnNames.indexOf("id");
  const noteColIndex = itemColumnNames.indexOf("optionalNote");

  function itemCell(grid: FakeSheetsService["grid"], sheetRowIndex: number) {
    const row = grid.sheet(itemGid).values()[sheetRowIndex] ?? [];
    return (colIndex: number): FakeCell | undefined =>
      row[expectedSheetLayout.startTableColIndex + colIndex];
  }

  it("fills each blank ID cell after the regen as its own send, listing every filled ID, and writes no other body cell", () => {
    const { grid, batchUpdateCount } = seedFixture({
      item: {
        bodyRows: [
          { id: itemRowId, optionalNote: "kept" },
          { optionalNote: "blank ID" },
        ],
      },
    });

    const { result } = runItem({ isSending: true });

    const filled = itemCell(grid, expectedSheetLayout.topDataRowIndex + 1);
    const filledId = String(filled(idColIndex));
    expect(filledId).toMatch(/^r:itm:.{7}$/);
    expect(filled(noteColIndex)).toBe("blank ID");
    expect(
      itemCell(grid, expectedSheetLayout.topDataRowIndex)(idColIndex),
    ).toBe(itemRowId);
    expect(batchUpdateCount()).toBe(1);
    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      plan: [
        {
          subject: { table: "item", column: "id" },
          what: `fill 1 blank ID: ${filledId}`,
        },
      ],
      landed: [
        {
          kind: "send",
          subject: { table: "item", column: "id" },
          what: `fill 1 blank ID: ${filledId}`,
        },
      ],
    });
  });

  it("lists the blank IDs in a dry run's plan and fills none", () => {
    const { grid, batchUpdateCount } = seedFixture({
      item: { bodyRows: [{}, {}] },
    });

    const { result } = runItem({ isSending: false });

    expect(batchUpdateCount()).toBe(0);
    expect(
      itemCell(grid, expectedSheetLayout.topDataRowIndex)(idColIndex) ?? null,
    ).toBeNull();
    expect(result.plan).toEqual([
      {
        subject: { table: "item", column: "id" },
        what: "fill 2 blank IDs",
        effects: [],
      },
    ]);
  });

  it("stops before any send when a Table with an ID column has no ID prefix, naming the Table", () => {
    const { batchUpdateCount } = seedFixture({ item: { bodyRows: [{}] } });
    vi.spyOn(
      TableConfigsOperator.prototype,
      "idPrefixesByTableId",
    ).mockReturnValue(new IdPrefixesByTableId(new Map()));

    expect(() => runItem({ isSending: true })).toThrow(
      'Table "Item" has an ID column but no ID prefix, so sync stopped before sending.',
    );
    expect(batchUpdateCount()).toBe(0);
  });

  it("sends a Fill Row IDs checkbox into the empty action-row cell above ID, listed in the plan, and the configs take it", () => {
    const { grid } = seedFixture({ item: { actionRowCells: {} } });

    const { result, files } = runItem({ isSending: true });

    expect(itemCell(grid, expectedSheetLayout.actionRowIndex)(idColIndex)).toBe(
      false,
    );
    expect(
      emittedItemColumn(Val.assert(files, "emitted files"), "id"),
    ).toContain('"actionRowCell": {"kind":"checkbox"}');
    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      plan: [
        {
          subject: { table: "item", column: "id" },
          what: "add the Fill Row IDs checkbox",
        },
      ],
    });
  });

  it("leaves an action-row cell holding text alone, with no conflict and no stop", () => {
    const { grid } = seedFixture({
      item: { actionRowCells: { id: "Keys" } },
    });

    const { result } = runItem({ isSending: true });

    expect(itemCell(grid, expectedSheetLayout.actionRowIndex)(idColIndex)).toBe(
      "Keys",
    );
    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      stops: [],
      plan: [],
    });
  });

  it("gives a Table whose ID is its first column the idFirst notice", () => {
    seedFixture({
      item: {
        columnNames: [
          "id",
          ...itemColumnNames.filter((columnName) => columnName !== "id"),
        ],
      },
    });

    const { result } = runItem({ isSending: false });

    expect(result.notices).toEqual([{ kind: "idFirst", table: "item" }]);
  });

  it("gives no notice when ID isn't the first column", () => {
    seedFixture();

    const { result } = runItem({ isSending: false });

    expect(idColIndex).toBeGreaterThan(0);
    expect(result.notices).toEqual([]);
  });
});

describe("SyncCoordinator send with no alterations", () => {
  it("returns every file's source together, reflecting the synced state", () => {
    seedFixture();

    const parsed = sendWithNoAlterations();
    expect(typeof parsed.tableConfigs).toBe("string");
    expect(typeof parsed.columnConfigs).toBe("string");
    expect(parsed.tableConfigs).toContain('"item"');
    expect(parsed.tableConfigs).toContain(
      "export const tableConfigs = makeTableConfigs({",
    );
    expect(parsed.tableConfigs).toContain(
      `"sheetGid": ${itemGid}, "idPrefix": "itm"`,
    );
    expect(typeof parsed.valueConfigs).toBe("string");
    expect(parsed.columnConfigs).toContain(itemColumns.optionalNote.columnId);
  });

  it("writes every make*Configs import from the framework's package name", () => {
    seedFixture();

    const parsed = sendWithNoAlterations();
    function importLine(maker: string): string {
      return `import { ${maker} } from "@byronbroughten/sheets-framework";`;
    }
    expect(parsed.tableConfigs).toContain(importLine("makeTableConfigs"));
    expect(parsed.tabConfigs).toContain(importLine("makeTabConfigs"));
    expect(parsed.conditionalFormatConfigs).toContain(
      importLine("makeConditionalFormatConfigs"),
    );
    expect(parsed.retiredIdPrefixes).toContain(
      importLine("makeRetiredIdPrefixes"),
    );
    expect(parsed.columnConfigs).toContain(importLine("makeColumnConfigs"));
    expect(parsed.valueConfigs).toContain(importLine("makeValueConfigs"));
  });

  it("emits the ID-name map as a type-only file", () => {
    seedFixture();

    const { idNameMap } = sendWithNoAlterations();
    expect(idNameMap).toContain("export interface IdNameMap {");
    expect(idNameMap).not.toMatch(/\b(const|let|function|import)\b/);
  });

  it("maps a Table with API access to its name, tab and columns", () => {
    seedFixture();

    const { tables } = parseIdNameMapFileSource(
      sendWithNoAlterations().idNameMap,
    );
    expect(tables.item).toEqual({
      name: "item",
      sheetGid: itemGid,
      columns: {
        optionalNote: itemColumns.optionalNote.columnId,
        requiredCount: itemColumns.requiredCount.columnId,
        id: itemColumns.id.columnId,
        name: itemColumns.name.columnId,
      },
    });
  });

  it("lists every tab by GID with its name, title and Tables", () => {
    seedFixture({ extraSheets: [notesSheet()] });

    const { tabs } = parseIdNameMapFileSource(
      sendWithNoAlterations().idNameMap,
    );
    expect(tabs[String(itemGid)]).toEqual({
      name: "item",
      title: "Item",
      tableIds: ["item"],
    });
    expect(tabs[String(notesGid)]).toEqual({
      name: "notes",
      title: "Notes",
      tableIds: ["notes"],
    });
  });

  it("emits no Spreadsheet Config file", () => {
    seedFixture();

    const parsed = sendWithNoAlterations();
    expect(Object.keys(parsed)).not.toContain("spreadsheetConfig");
    expect(Object.values(parsed).join("\n")).not.toContain(
      "makeSpreadsheetConfig",
    );
  });

  it("carries the untyped-column summary back, since no run status cell will show it", () => {
    seedFixture();

    const { result } = runWithNoAlterations();

    expect(result.configs.notes).toEqual([
      expect.stringContaining("column(s) across 1 Table(s)"),
    ]);
  });
});

describe("SyncCoordinator send with no alterations API access", () => {
  const logGid = getTableTraitByName("log", "sheetGid");

  function logSheet(): FakeSheetProperties {
    return fakeTableSheet.build({
      sheetId: logGid,
      title: "Log",
      tableId: "log",
      name: "Log",
      columnConfigs: columnConfigs.log,
      columnNames: ["entry", "amount"],
      bodyRows: [{}],
    });
  }

  function tableConfigKeys(files: ConfigRegeneration): string[] {
    return [...files.tableConfigs.matchAll(/^ {2}"([^"]+)": \{/gm)].map(
      (match) => Val.assert(match[1], "Table key"),
    );
  }

  it("gives API access to each live Table the previous configs hold, in their order", () => {
    stubSheetsService({ sheets: [logSheet(), notesSheet(), itemSheet()] });

    expect(tableConfigKeys(sendWithNoAlterations())).toEqual(["item", "log"]);
  });

  it("gives a Table the previous configs don't hold no prefix, no column IDs and no columns in the ID-name map", () => {
    const { grid } = seedFixture({ extraSheets: [notesSheet()] });

    const files = sendWithNoAlterations();

    expect(tableConfigKeys(files)).toEqual(["item"]);
    expect(
      grid.sheet(notesGid).values()[expectedSheetLayout.colIdRowIndex],
    ).toEqual([""]);
    expect(parseIdNameMapFileSource(files.idNameMap).tables.notes).toEqual({
      name: "deskNotes",
      sheetGid: notesGid,
    });
  });
});

describe("SyncCoordinator send with no alterations Empty value allowed", () => {
  function generateItemColumns() {
    const coordinator = SyncCoordinator.init();
    sendWithNoAlterations(coordinator);
    return Val.assert(
      coordinator.columnConfigOperator.newColumnConfigs().item,
      "Item column configs",
    );
  }

  it("carries each column's Empty value allowed forward by column ID, through a header rename", () => {
    const optionalNote: ItemColumnName = "optionalNote";
    seedFixture({
      item: {
        columns: {
          [optionalNote]: {
            columnId: itemColumns.optionalNote.columnId,
            header: "Side note",
          },
        },
      },
    });

    const item = generateItemColumns();

    expect(itemColumns.optionalNote.emptyValueAllowed).toBe(true);
    expect(itemColumns.requiredCount.emptyValueAllowed).toBe(false);
    expect(item.sideNote?.emptyValueAllowed).toBe(true);
    expect(item.requiredCount?.emptyValueAllowed).toBe(false);
  });

  it("gives a new column Empty value allowed false", () => {
    seedFixture({
      item: { columns: { extra: { columnId: "", header: "Extra" } } },
    });

    const item = generateItemColumns();

    expect(item.extra?.columnId).toMatch(/^c:itm:/);
    expect(item.extra?.emptyValueAllowed).toBe(false);
  });
});

describe("SyncCoordinator send with no alterations ID-prefix guard", () => {
  it("stops before writing when a Table's live ID prefix differs from the configs", () => {
    const zzzColumns = Object.fromEntries(
      Object.entries(itemColumns).map(([columnName, column]) => [
        columnName,
        {
          header: column.header,
          columnId: column.columnId.replace("c:itm:", "c:zzz:"),
        },
      ]),
    );
    const { batchUpdateCount } = seedFixture({
      item: { columns: zzzColumns },
    });

    expect(() => sendWithNoAlterations()).toThrow(
      'The regen would change an ID prefix, so it stopped before writing. Table "item" had ID prefix "itm" and is now "zzz".',
    );
    expect(batchUpdateCount()).toBe(0);
  });
});

describe("SyncCoordinator send with no alterations ID prefix", () => {
  const widgetGid = 888001;
  // A Table the previous configs grant, standing in for the fixture's own.
  const widgetTableId = getTableTraitByName("log", "tableId");
  const widgetIdPrefix = getTableTraitByName("log", "idPrefix");

  function widgetSheet(options: {
    title: string;
    tableName?: string;
    columnIds?: readonly string[];
    headers?: readonly string[];
  }): FakeSheetProperties {
    const headers = options.headers ?? ["Name"];
    return fakeTableSheet.build({
      sheetId: widgetGid,
      title: options.title,
      tableId: widgetTableId,
      name: options.tableName ?? options.title,
      columnConfigs: columnsByHeader(headers, options.columnIds),
      columnNames: headers,
      bodyRows: [{}],
    });
  }

  it("samples an existing sheet's prefix from its column IDs, whatever the tab title is now", () => {
    seedFixture({
      extraSheets: [
        widgetSheet({
          title: "Renamed Tab",
          columnIds: [`c:${widgetIdPrefix}:abc1234`],
        }),
      ],
    });

    const parsed = sendWithNoAlterations();
    expect(parsed.tableConfigs).toContain(
      tableConfigsEntry({
        tableKey: "renamedTab",
        tableId: widgetTableId,
        tableName: "Renamed Tab",
        sheetGid: widgetGid,
        idPrefix: widgetIdPrefix,
        hasIdColumn: false,
      }),
    );
  });

  it("mints blank column ID cells on an existing sheet with its prefix", () => {
    seedFixture({
      extraSheets: [
        widgetSheet({
          title: "Widget",
          columnIds: [`c:${widgetIdPrefix}:abc1234`, ""],
          headers: ["Name", "Notes"],
        }),
      ],
    });

    const parsed = sendWithNoAlterations();
    const minted =
      parsed.columnConfigs.match(
        new RegExp(`c:${widgetIdPrefix}:[^"]+`, "g"),
      ) ?? [];
    expect(minted).toHaveLength(2);
    expect(minted).toContain(`c:${widgetIdPrefix}:abc1234`);
  });

  it("stops when one sheet's column IDs carry mixed prefixes, naming the sheet and stray IDs", () => {
    seedFixture({
      extraSheets: [
        widgetSheet({
          title: "Widget",
          columnIds: [`c:${widgetIdPrefix}:abc1234`, "c:gzm:xyz1234"],
          headers: ["Name", "Notes"],
        }),
      ],
    });

    expect(() => sendWithNoAlterations()).toThrow(/Widget.*c:gzm:xyz1234/);
  });

  it("fails when two Tables' names give the same key, naming both", () => {
    seedFixture({
      extraSheets: [widgetSheet({ title: "Widget", tableName: "item" })],
    });

    expect(() => sendWithNoAlterations()).toThrow(
      'Tables "Item" and "item" both give the key "item".',
    );
  });
});

describe("SyncCoordinator send with no alterations head-row overlap", () => {
  const recordsGid = 888000222;

  interface RecordsTable {
    name: string;
    headerRowIndex: number;
    startColumnIndex: number;
    endColumnIndex: number;
    endRowIndex: number;
    isManaged: boolean;
  }

  // A managed Table takes a Table ID the previous configs grant.
  const grantedTableIdsByName: Record<string, string> = {
    Leases: getTableTraitByName("log", "tableId"),
    Tenants: getTableTraitByName("runItem", "tableId"),
    Units: getTableTraitByName("computed", "tableId"),
    Owners: getTableTraitByName("dates", "tableId"),
  };

  function recordsTableId(table: RecordsTable): string {
    if (!table.isManaged) return `records-${table.name.toLowerCase()}`;
    return Val.assert(
      grantedTableIdsByName[table.name],
      "granted Table ID of a managed Records Table",
    );
  }

  // Several Tables on one sheet, which the one-Table builder can't place.
  function recordsSheet(tables: readonly RecordsTable[]): FakeSheetProperties {
    const rows: Record<number, FakeCellValue[]> = {};
    tables.forEach((table) => {
      const headerRow = (rows[table.headerRowIndex] ??= []);
      const topRow = (rows[table.headerRowIndex + 1] ??= []);
      for (
        let col = table.startColumnIndex;
        col < table.endColumnIndex;
        col++
      ) {
        headerRow[col] = `${table.name} ${col}`;
        topRow[col] ??= "";
      }
    });
    return {
      sheetId: recordsGid,
      title: "Records",
      rows: buildGridRows(
        Object.fromEntries(
          Object.entries(rows).map(([rowIndex, row]) => [
            rowIndex,
            Array.from(row, (cell) => cell ?? ""),
          ]),
        ),
      ),
      tables: tables.map((table) => ({
        tableId: recordsTableId(table),
        name: table.name,
        startRowIndex: table.headerRowIndex,
        startColumnIndex: table.startColumnIndex,
        endRowIndex: table.endRowIndex,
        endColumnIndex: table.endColumnIndex,
      })),
    };
  }

  function seedRecords(tables: readonly RecordsTable[]) {
    return seedFixture({ extraSheets: [recordsSheet(tables)] });
  }

  // The managed Tables' blank column ID cells are what a run that went on would write.
  function expectRefused(
    tables: readonly RecordsTable[],
    ...messages: string[]
  ): void {
    const { grid } = seedRecords(tables);
    const recordsValues = () => grid.sheet(recordsGid).values();
    const before = recordsValues();

    let parsed: ConfigRegeneration | undefined;
    expect(() => {
      parsed = sendWithNoAlterations();
    }).toThrow(messages.join(" "));
    expect(parsed).toBeUndefined();
    expect(recordsValues()).toEqual(before);
  }

  // Rows 4–9 are its body, 0-based.
  const leases: RecordsTable = {
    name: "Leases",
    headerRowIndex: 3,
    startColumnIndex: 1,
    endColumnIndex: 4,
    endRowIndex: 10,
    isManaged: true,
  };

  it("refuses a Table stacked under another and starting left of it, its head rows on the other's last body row", () => {
    expectRefused(
      [
        leases,
        {
          name: "Tenants",
          headerRowIndex: 12,
          startColumnIndex: 0,
          endColumnIndex: 3,
          endRowIndex: 14,
          isManaged: true,
        },
      ],
      `Table "Tenants" on sheet "Records" has head rows in rows 10–12 that sit on Table "Leases". Move "Tenants" so that the 3 rows above its header sit clear of "Leases".`,
    );
  });

  it("refuses a Table whose head rows sit on another managed Table's head rows", () => {
    expectRefused(
      [
        {
          name: "Leases",
          headerRowIndex: 3,
          startColumnIndex: 0,
          endColumnIndex: 3,
          endRowIndex: 5,
          isManaged: true,
        },
        {
          name: "Tenants",
          headerRowIndex: 5,
          startColumnIndex: 0,
          endColumnIndex: 3,
          endRowIndex: 7,
          isManaged: true,
        },
      ],
      `Table "Tenants" on sheet "Records" has head rows in rows 3–5 that sit on Table "Leases". Move "Tenants" so that the 3 rows above its header sit clear of "Leases".`,
    );
  });

  it("refuses a managed Table whose head rows sit on a Table the app doesn't manage", () => {
    expectRefused(
      [
        { ...leases, isManaged: false },
        {
          name: "Tenants",
          headerRowIndex: 12,
          startColumnIndex: 0,
          endColumnIndex: 3,
          endRowIndex: 14,
          isManaged: true,
        },
      ],
      `Table "Tenants" on sheet "Records" has head rows in rows 10–12 that sit on Table "Leases". Move "Tenants" so that the 3 rows above its header sit clear of "Leases".`,
    );
  });

  it("lists two overlaps on one sheet in one message", () => {
    expectRefused(
      [
        {
          name: "Leases",
          headerRowIndex: 3,
          startColumnIndex: 0,
          endColumnIndex: 2,
          endRowIndex: 10,
          isManaged: true,
        },
        {
          name: "Units",
          headerRowIndex: 3,
          startColumnIndex: 3,
          endColumnIndex: 5,
          endRowIndex: 10,
          isManaged: true,
        },
        {
          name: "Tenants",
          headerRowIndex: 12,
          startColumnIndex: 0,
          endColumnIndex: 2,
          endRowIndex: 14,
          isManaged: true,
        },
        {
          name: "Owners",
          headerRowIndex: 11,
          startColumnIndex: 3,
          endColumnIndex: 5,
          endRowIndex: 13,
          isManaged: true,
        },
      ],
      `Table "Tenants" on sheet "Records" has head rows in rows 10–12 that sit on Table "Leases". Move "Tenants" so that the 3 rows above its header sit clear of "Leases".`,
      `Table "Owners" on sheet "Records" has head rows in rows 9–11 that sit on Table "Units". Move "Owners" so that the 3 rows above its header sit clear of "Units".`,
    );
  });

  it("refuses a granted Table stacked clear of another but below the header zone, naming it", () => {
    seedRecords([
      leases,
      {
        name: "Tenants",
        headerRowIndex: 13,
        startColumnIndex: 0,
        endColumnIndex: 3,
        endRowIndex: 15,
        isManaged: true,
      },
    ]);

    expect(() => sendWithNoAlterations()).toThrowError(
      /^Table "Tenants" on "Records" \(gid \d+\) must have its header row on row 4\.$/,
    );
  });

  it("refuses a recorded Table moved below the header zone with generation's own fix", () => {
    seedFixture({
      item: { startRowIndex: expectedSheetLayout.tableHeaderRowIndex + 2 },
    });

    expect(() => sendWithNoAlterations()).toThrowError(
      /^Table "Item" on "Item" \(gid \d+\) must have its header row on row 4\.$/,
    );
  });

  it("leaves an ungranted Table below the header zone to the operator", () => {
    seedRecords([
      leases,
      {
        name: "Tenants",
        headerRowIndex: 13,
        startColumnIndex: 0,
        endColumnIndex: 3,
        endRowIndex: 15,
        isManaged: false,
      },
    ]);

    const parsed = sendWithNoAlterations();
    expect(parsed.tableConfigs).toContain('"leases"');
    expect(parsed.tableConfigs).not.toContain('"tenants"');
  });

  it("allows two Tables side by side, touching", () => {
    seedRecords([
      {
        name: "Leases",
        headerRowIndex: 3,
        startColumnIndex: 0,
        endColumnIndex: 2,
        endRowIndex: 6,
        isManaged: true,
      },
      {
        name: "Tenants",
        headerRowIndex: 3,
        startColumnIndex: 2,
        endColumnIndex: 4,
        endRowIndex: 5,
        isManaged: true,
      },
    ]);

    const parsed = sendWithNoAlterations();
    expect(parsed.tableConfigs).toContain('"tenants"');
    expect(parsed.tableConfigs).toContain('"leases"');
  });
});

describe("SyncCoordinator conditional-format rule read", () => {
  const rulesGid = 777000222;
  const widgetGid = 777000333;
  // A Table the previous configs grant, so its rules go under its Table key.
  const widgetTableId = getTableTraitByName("log", "tableId");
  const separateRuleGet = "sheets(properties(sheetId),conditionalFormats)";

  function backgroundRule(
    range: GridRange,
  ): GoogleAppsScript.Sheets.Schema.ConditionalFormatRule {
    return {
      ranges: [range],
      booleanRule: {
        condition: {
          type: "CUSTOM_FORMULA",
          values: [{ userEnteredValue: "=$B5>1" }],
        },
        format: { backgroundColor: { red: 1, green: 0, blue: 0 } },
      },
    };
  }

  function rulesSheet(): FakeSheetProperties {
    return {
      sheetId: rulesGid,
      title: "Rules",
      rows: buildGridRows({}),
      conditionalFormats: [
        backgroundRule({
          sheetId: rulesGid,
          startRowIndex: 4,
          endRowIndex: 8,
          startColumnIndex: 1,
          endColumnIndex: 2,
        }),
      ],
    };
  }

  it("brings every sheet's rules into Raw on the regen's plain get, with no separate rule get", () => {
    const service = seedFixture({ extraSheets: [rulesSheet()] });
    const coordinator = SyncCoordinator.init();

    sendWithNoAlterations(coordinator);

    expect(coordinator.ss.raw.sheet(rulesGid).conditionalFormatRules()).toEqual(
      [
        {
          kind: "boolean",
          ranges: [
            {
              sheetId: rulesGid,
              startRowIndex: 4,
              endRowIndex: 8,
              startColumnIndex: 1,
              endColumnIndex: 2,
            },
          ],
          condition: { type: "CUSTOM_FORMULA", formula: "=$B5>1" },
          format: { backgroundColor: { red: 1, green: 0, blue: 0 } },
        },
      ],
    );
    expect(service.getCalls.map(({ fields }) => fields)).not.toContain(
      separateRuleGet,
    );
  });

  it("writes a Table's rules under its Table key and a tab's under its tab name, each by its rule handle", () => {
    seedFixture({
      extraSheets: [
        rulesSheet(),
        {
          ...fakeTableSheet.build({
            sheetId: widgetGid,
            title: "Widget",
            tableId: widgetTableId,
            name: "Widget",
            columnConfigs: columnsByHeader(["Name", "Score"]),
            columnNames: ["Name", "Score"],
            bodyRows: [{}],
          }),
          conditionalFormats: [
            backgroundRule({
              sheetId: widgetGid,
              startRowIndex: 4,
              endRowIndex: 5,
              startColumnIndex: 1,
              endColumnIndex: 2,
            }),
          ],
        },
      ],
    });

    const { conditionalFormatConfigs } = sendWithNoAlterations();

    expect(conditionalFormatConfigs).toContain(
      "export const conditionalFormatConfigs = makeConditionalFormatConfigs({",
    );
    expect(conditionalFormatConfigs).toContain(
      `  "tables": {\n    "widget": {\n      "score | =\${@score}>1 | background #ff0000": { "kind": "boolean", "tableId": "${widgetTableId}", "columns": ["score"], "condition": {"type":"CUSTOM_FORMULA","formula":"=\${@score}>1"}, "format": {"backgroundColor":"#ff0000"} }\n    }\n  },`,
    );
    expect(conditionalFormatConfigs).toContain(
      `  "tabs": {\n    "rules": {\n      "B5:B8 | =$B5>1 | background #ff0000": { "kind": "boolean", "sheetGid": ${rulesGid}, "ranges": ["B5:B8"], "condition": {"type":"CUSTOM_FORMULA","formula":"=$B5>1"}, "format": {"backgroundColor":"#ff0000"} }\n    }\n  }\n});`,
    );
  });
});

describe("SyncCoordinator send with no alterations value sources", () => {
  const listGid = 889001;
  const ledgerGid = 889002;
  // The previous configs grant API access, so an API-access fixture takes a Table ID they hold.
  const listTableId = getTableTraitByName("log", "tableId");
  const ledgerTableId = getTableTraitByName("runItem", "tableId");
  const listPaymentTableId = getTableTraitByName("computed", "tableId");

  interface FixtureColumn {
    header: string;
    values?: readonly FakeCell[];
    columnType?: string;
    validationValues?: string[];
    validationConditionType?: string;
    actionRowCell?: FakeCellValue;
    groupHeading1?: string;
  }

  // Omit tableId for a Table the previous configs don't hold.
  function tableSheet(props: {
    sheetId: number;
    title: string;
    tableId?: string;
    columns: readonly FixtureColumn[];
  }): FakeSheetProperties {
    const { columns } = props;
    const headers = columns.map((column) => column.header);
    const rowCount = Math.max(
      1,
      ...columns.map((column) => column.values?.length ?? 0),
    );
    const bodyRows: FakeBodyRow<string>[] = Array.from(
      { length: rowCount },
      (_, rowOffset) =>
        Object.fromEntries(
          columns.map((column) => [
            column.header,
            column.values?.[rowOffset] ?? null,
          ]),
        ),
    );
    return fakeTableSheet.build({
      sheetId: props.sheetId,
      title: props.title,
      tableId: props.tableId,
      name: props.title,
      columnConfigs: columnsByHeader(headers),
      columnNames: headers,
      bodyRows,
      headRows: {
        1: columns.map((column) => column.actionRowCell ?? null),
        2: columns.map((column) => column.groupHeading1 ?? null),
      },
      columnTypes: byColIndex(columns, (column) => column.columnType),
      columnValidationValues: byColIndex(
        columns,
        (column) => column.validationValues,
      ),
      columnValidationConditionTypes: byColIndex(
        columns,
        (column) => column.validationConditionType,
      ),
    });
  }

  function byColIndex<V>(
    columns: readonly FixtureColumn[],
    read: (column: FixtureColumn) => V | undefined,
  ): Record<number, V> {
    return columns.reduce<Record<number, V>>((byIndex, column, colIndex) => {
      const value = read(column);
      if (value !== undefined) byIndex[colIndex] = value;
      return byIndex;
    }, {});
  }

  function ledgerSheet(columns: readonly FixtureColumn[]): FakeSheetProperties {
    return tableSheet({
      sheetId: ledgerGid,
      title: "Ledger",
      tableId: ledgerTableId,
      columns,
    });
  }

  function listSheet(columns: readonly FixtureColumn[]): FakeSheetProperties {
    return tableSheet({
      sheetId: listGid,
      title: "List",
      tableId: listTableId,
      columns,
    });
  }

  function generate() {
    const coordinator = SyncCoordinator.init();
    const files = sendWithNoAlterations(coordinator);
    return {
      valueConfigs: emittedValueConfigs(files),
      ledgerValueName: (columnName: string) =>
        coordinator.columnConfigOperator.newColumnConfigs().ledger?.[columnName]
          ?.valueName,
    };
  }

  // A source fetch reads one whole column; row samples span the Table.
  function fullyFetchedListColumnIndexes(service: FakeSheetsService): number[] {
    const indexes = service.getByDataFilterCalls.flatMap((resource) =>
      (
        (resource as { dataFilters?: { gridRange?: GridRange }[] })
          .dataFilters ?? []
      )
        .map(({ gridRange }) => gridRange)
        .filter(
          (range): range is GridRange =>
            range?.sheetId === listGid &&
            range.startRowIndex === 0 &&
            (range.endColumnIndex ?? 0) - (range.startColumnIndex ?? 0) === 1,
        )
        .map((range) => Val.assert(range.startColumnIndex, "start column")),
    );
    return [...new Set(indexes)].sort();
  }

  function emittedValueConfigs(files: ConfigRegeneration): unknown {
    const match = files.valueConfigs.match(/makeValueConfigs\(([\s\S]*)\);/);
    return JSON.parse(Val.assert(match?.[1], "emitted value configs"));
  }

  it("types a dropdown from a Let api access Table column as that column's fixed list", () => {
    seedFixture({
      extraSheets: [
        listSheet([
          { header: "Payment status", values: ["Paid", "Unpaid", "Late"] },
        ]),
        ledgerSheet([
          {
            header: "Status",
            values: ["Paid"],
            validationValues: ["=List[Payment status]"],
          },
        ]),
      ],
    });

    const { valueConfigs, ledgerValueName } = generate();

    expect(ledgerValueName("status")).toBe("listPaymentStatus");
    expect(valueConfigs).toEqual({
      listPaymentStatus: ["Paid", "Unpaid", "Late"],
    });
  });

  function seedNameAndIdSources(): FakeSheetsService {
    return seedFixture({
      extraSheets: [
        listSheet([
          { header: "ID", values: ["r:lst:a1"] },
          { header: "Name", values: ["Unit 2B"], columnType: "DOUBLE" },
          { header: "Kind", values: ["Flat"] },
        ]),
        ledgerSheet([
          { header: "Unit", validationValues: ["=List[Name]"] },
          { header: "Unit row", validationValues: ["=List[ID]"] },
          { header: "Kind", validationValues: ["=List[Kind]"] },
        ]),
      ],
    });
  }

  it("emits a dropdown from a Table's Name or ID column as a live reference, whatever its type", () => {
    seedNameAndIdSources();

    const { valueConfigs, ledgerValueName } = generate();

    expect(
      ["unit", "unitRow", "kind"].map((columnName) =>
        ledgerValueName(columnName),
      ),
    ).toEqual(["listName", "listId", "listKind"]);
    expect(valueConfigs).toEqual({
      listName: { kind: "liveReference" },
      listId: { kind: "liveReference" },
      listKind: ["Flat"],
    });
  });

  it("fetches no Name source column, only the fixed-list ones and the ID column Fill Row IDs reads", () => {
    const service = seedNameAndIdSources();

    generate();

    // Column 0 is the ID column, read after the regen to fill its blanks.
    expect(fullyFetchedListColumnIndexes(service)).toEqual([0, 2]);
  });

  it("keeps a one-word camel-case header's casing in the value name", () => {
    seedFixture({
      extraSheets: [
        listSheet([{ header: "CapEx", values: ["Roof"] }]),
        ledgerSheet([{ header: "CapEx", validationValues: ["=List[CapEx]"] }]),
      ],
    });

    const { valueConfigs, ledgerValueName } = generate();

    expect(ledgerValueName("capEx")).toBe("listCapEx");
    expect(valueConfigs).toEqual({ listCapEx: ["Roof"] });
  });

  it("types every member by its source's declared type, numbers for number and date", () => {
    seedFixture({
      extraSheets: [
        listSheet([
          { header: "Rent", values: [950, 1200], columnType: "DOUBLE" },
          { header: "Due date", values: [45000, 45031], columnType: "DATE" },
          { header: "Unit", values: ["2B", 101], columnType: "TEXT" },
        ]),
        ledgerSheet([
          { header: "Rent", validationValues: ["=List[Rent]"] },
          { header: "Due", validationValues: ["=List[Due date]"] },
          { header: "Unit", validationValues: ["=List[Unit]"] },
        ]),
      ],
    });

    const { valueConfigs } = generate();

    expect(valueConfigs).toEqual({
      listRent: [950, 1200],
      listDueDate: [45000, 45031],
      listUnit: ["2B", "101"],
    });
  });

  it("fails the config sync when a declared source holds a member of another primitive, naming the source", () => {
    seedFixture({
      extraSheets: [
        listSheet([
          { header: "Rent", values: [950, "TBD"], columnType: "DOUBLE" },
        ]),
        ledgerSheet([{ header: "Rent", validationValues: ["=List[Rent]"] }]),
      ],
    });

    expect(() => sendWithNoAlterations()).toThrow(
      'Value source "List[Rent]" declares number members, but holds "TBD".',
    );
  });

  it("keeps each member's own primitive on an untyped source", () => {
    seedFixture({
      extraSheets: [
        listSheet([{ header: "Score", values: ["N/A", 3, true] }]),
        ledgerSheet([{ header: "Score", validationValues: ["=List[Score]"] }]),
      ],
    });

    const { valueConfigs } = generate();

    expect(valueConfigs).toEqual({ listScore: ["N/A", 3, true] });
  });

  it("drops blanks and collapses duplicates to their first position, keeping sheet order", () => {
    seedFixture({
      extraSheets: [
        listSheet([
          {
            header: "Payment status",
            values: ["Unpaid", null, "Paid", "Unpaid", "", "Late", "Paid"],
          },
        ]),
        ledgerSheet([
          { header: "Status", validationValues: ["=List[Payment status]"] },
        ]),
      ],
    });

    const { valueConfigs } = generate();

    expect(valueConfigs).toEqual({
      listPaymentStatus: ["Unpaid", "Paid", "Late"],
    });
  });

  it("gives every dropdown on one source the one value name", () => {
    seedFixture({
      extraSheets: [
        listSheet([{ header: "Payment status", values: ["Paid", "Unpaid"] }]),
        ledgerSheet([
          { header: "Status", validationValues: ["=List[Payment status]"] },
          {
            header: "Deposit status",
            validationValues: ["=List[Payment status]"],
          },
        ]),
      ],
    });

    const { valueConfigs, ledgerValueName } = generate();

    expect(ledgerValueName("status")).toBe("listPaymentStatus");
    expect(ledgerValueName("depositStatus")).toBe("listPaymentStatus");
    expect(valueConfigs).toEqual({ listPaymentStatus: ["Paid", "Unpaid"] });
  });

  it("leaves a dropdown untyped, and in the summary, unless it names a column of a Let api access Table", () => {
    seedFixture({
      extraSheets: [
        listSheet([{ header: "Payment status", values: ["Paid", "Unpaid"] }]),
        ledgerSheet([
          {
            header: "Non-api",
            values: ["Paid"],
            validationValues: ["=Hidden[Payment status]"],
          },
          {
            header: "Unknown header",
            values: ["Paid"],
            validationValues: ["=List[Nope]"],
          },
          {
            header: "A1 range",
            values: ["Paid"],
            validationValues: ["=List!$A$5:$A$9"],
          },
          {
            header: "Item list",
            values: ["Paid"],
            validationValues: ["Paid", "Unpaid"],
          },
        ]),
        tableSheet({
          sheetId: 889003,
          title: "Hidden",
          columns: [{ header: "Payment status", values: ["Paid"] }],
        }),
      ],
    });
    const logger = stubLogger();

    const coordinator = SyncCoordinator.init();
    const { result, files } = runWithNoAlterations(coordinator);

    expect(emittedValueConfigs(Val.assert(files, "emitted files"))).toEqual({});
    const ledger = coordinator.columnConfigOperator.newColumnConfigs().ledger;
    expect(
      ["nonApi", "unknownHeader", "a1Range", "itemList"].map(
        (columnName) => ledger?.[columnName]?.valueName,
      ),
    ).toEqual(["string", "string", "string", "string"]);
    expect(result.configs.notes).toEqual([
      expect.stringContaining("column(s) across"),
    ]);
    expect(logger.log).toHaveBeenCalledWith(
      'Untyped columns on Table "Ledger" (4): Non-api, Unknown header, A1 range, Item list',
    );
  });

  it("fails the config sync when a source's value name is a framework value name, naming both", () => {
    seedFixture({
      extraSheets: [
        tableSheet({
          sheetId: listGid,
          title: "Number",
          tableId: listTableId,
          columns: [{ header: "#", values: [1, 2] }],
        }),
        ledgerSheet([{ header: "Pick", validationValues: ["=Number[#]"] }]),
      ],
    });

    expect(() => sendWithNoAlterations()).toThrow(
      'Value source "Number[#]" takes the value name "number", which is already the framework value name "number".',
    );
  });

  it("fails the config sync when two sources take one value name, naming both", () => {
    seedFixture({
      extraSheets: [
        listSheet([{ header: "Payment status", values: ["Paid"] }]),
        tableSheet({
          sheetId: 889004,
          title: "List Payment",
          tableId: listPaymentTableId,
          columns: [{ header: "Status", values: ["Open"] }],
        }),
        ledgerSheet([
          { header: "Status", validationValues: ["=List[Payment status]"] },
          { header: "State", validationValues: ["=List Payment[Status]"] },
        ]),
      ],
    });

    expect(() => sendWithNoAlterations()).toThrow(
      'Value sources "List[Payment status]" and "List Payment[Status]" both take the value name "listPaymentStatus".',
    );
  });

  describe("the baseline each column records", () => {
    function generateLedgerColumns() {
      const coordinator = SyncCoordinator.init();
      const files = sendWithNoAlterations(coordinator);
      return {
        files,
        ledger: Val.assert(
          coordinator.columnConfigOperator.newColumnConfigs().ledger,
          "Ledger column configs",
        ),
      };
    }

    it("records formula text in brace form, naming its own Table's columns bare, another Table's in full, and this row with @", () => {
      seedFixture({
        extraSheets: [
          listSheet([{ header: "Payment status", values: ["Paid"] }]),
          ledgerSheet([
            { header: "Amount", values: [10] },
            {
              header: "Double",
              values: [
                {
                  value: 20,
                  isFormula: true,
                  formula: "=SINGLE(Ledger[Amount])*2",
                },
              ],
            },
            {
              header: "Statuses",
              values: [
                {
                  value: 1,
                  isFormula: true,
                  formula: '=COUNTA(List[Payment status])&"Ledger[Amount]"',
                },
              ],
            },
          ]),
        ],
      });

      const { files, ledger } = generateLedgerColumns();

      expect(ledger.amount?.formula).toBeNull();
      expect(ledger.double?.formula).toBe("=${@amount}*2");
      expect(ledger.statuses?.formula).toBe(
        '=COUNTA(${list[paymentStatus]})&"Ledger[Amount]"',
      );
      expect(files.columnConfigs).toContain('"formula": "=${@amount}*2"');
    });

    it("records the column type, with the rule a dropdown carries", () => {
      seedFixture({
        extraSheets: [
          listSheet([{ header: "Payment status", values: ["Paid"] }]),
          ledgerSheet([
            { header: "Amount", values: [10], columnType: "CURRENCY" },
            {
              header: "Status",
              values: ["Paid"],
              columnType: "DROPDOWN",
              validationConditionType: "ONE_OF_RANGE",
              validationValues: ["=List[Payment status]"],
            },
            { header: "Note", values: ["x"] },
          ]),
        ],
      });

      const { ledger } = generateLedgerColumns();

      expect(ledger.amount?.columnType).toEqual({
        type: "CURRENCY",
        rule: null,
      });
      expect(ledger.status?.columnType).toEqual({
        type: "DROPDOWN",
        rule: { condition: "ONE_OF_RANGE", values: ["=List[Payment status]"] },
      });
      expect(ledger.note?.columnType).toEqual({ type: null, rule: null });
    });

    it("records each action-row cell as empty, a checkbox or a level-2 group heading", () => {
      seedFixture({
        extraSheets: [
          ledgerSheet([
            { header: "Amount", values: [10] },
            { header: "Run", values: [false], actionRowCell: false },
            { header: "Due", values: [1], actionRowCell: "Dates" },
          ]),
        ],
      });

      const { ledger } = generateLedgerColumns();

      expect(ledger.amount?.actionRowCell).toBeNull();
      expect(ledger.run?.actionRowCell).toEqual({ kind: "checkbox" });
      expect(ledger.due?.actionRowCell).toEqual({
        kind: "groupHeading2",
        text: "Dates",
      });
    });

    it("records a spanning level-1 group heading on its first cell only", () => {
      seedFixture({
        extraSheets: [
          ledgerSheet([
            { header: "Amount", values: [10], groupHeading1: "Money" },
            { header: "Fee", values: [1] },
          ]),
        ],
      });

      const { ledger } = generateLedgerColumns();

      expect(ledger.amount?.groupHeading1).toBe("Money");
      expect(ledger.fee?.groupHeading1).toBeNull();
    });
  });
});
