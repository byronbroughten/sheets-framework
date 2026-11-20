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
import {
  bodyRowsWithNote,
  type ItemColumnName,
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

describe("SyncCoordinator.run column moves", () => {
  const columnNames = Object.keys(itemColumns) as ItemColumnName[];
  const headers = columnNames.map((name) => itemColumns[name].header);

  function nameAt(colIndex: number): ItemColumnName {
    return Val.assert(columnNames[colIndex], `Item column ${colIndex}`);
  }

  function headerAt(colIndex: number): string {
    return Val.assert(headers[colIndex], `Item header ${colIndex}`);
  }

  function itemSheetRows(service: FakeSheetsService): FakeCellValue[][] {
    return service.grid.sheet(itemGid).values();
  }

  function headersOf(rows: FakeCellValue[][]): FakeCellValue[] {
    const headerRow = rows[expectedSheetLayout.tableHeaderRowIndex] ?? [];
    return headerRow.filter((cell) => cell !== "");
  }

  function cellUnder(
    rows: FakeCellValue[][],
    header: string,
    rowOffset: number,
  ): FakeCellValue | undefined {
    const { tableHeaderRowIndex, colIdRowIndex } = expectedSheetLayout;
    const colIndex = rows[tableHeaderRowIndex]?.indexOf(header) ?? -1;
    if (rowOffset < 0) return rows[colIdRowIndex]?.[colIndex];
    return rows[tableHeaderRowIndex + 1 + rowOffset]?.[colIndex];
  }

  function itemTableColumns(service: FakeSheetsService): number {
    const range = Val.assert(
      service.grid.sheet(itemGid).tables[0]?.range,
      "Item Table range",
    );
    return (
      Val.assert(range.endColumnIndex, "Item Table end") -
      Val.assert(range.startColumnIndex, "Item Table start")
    );
  }

  function moveAlterations(
    column: ItemColumnName,
    before: ItemColumnName,
  ): PendingAlterations {
    return { tables: { item: { columns: { [column]: { before } } } } };
  }

  it("moves a column to before one in the middle, keeping its column ID and data", () => {
    const service = seedFixture({ item: { bodyRows: bodyRowsWithNote } });
    const noteHeader = itemColumns.optionalNote.header;
    const tableColumns = itemTableColumns(service);
    const moved = 2;
    const movedHeader = headerAt(moved);

    const { result, files } = runItem({
      isSending: true,
      alterations: moveAlterations(nameAt(moved), nameAt(1)),
    });

    const after = itemSheetRows(service);
    const expected = [...headers];
    expected.splice(moved, 1);
    expected.splice(1, 0, movedHeader);
    expect(headersOf(after)).toEqual(expected);
    expect(itemTableColumns(service)).toBe(tableColumns);
    expect(cellUnder(after, movedHeader, -1)).toBe(
      itemColumns[nameAt(moved)].columnId,
    );
    expect(cellUnder(after, noteHeader, 1)).toBe("kept");
    expect(files?.alterations).toContain("makePendingAlterations({})");
    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      plan: [
        {
          subject: { table: "item", column: nameAt(moved), path: "before" },
          what: `set to ${nameAt(1)}`,
        },
      ],
      landed: [
        { subject: { table: "item", column: nameAt(moved), path: "before" } },
      ],
    });
  });

  it("moves a column to the first position as the columns before it moving right, inside the Table", () => {
    const service = seedFixture();
    const tableColumns = itemTableColumns(service);

    const { result } = runItem({
      isSending: true,
      alterations: moveAlterations(nameAt(2), nameAt(0)),
    });

    const [first, second, third, ...rest] = headers;
    expect(headersOf(itemSheetRows(service))).toEqual([
      third,
      first,
      second,
      ...rest,
    ]);
    expect(itemTableColumns(service)).toBe(tableColumns);
    expect(service.grid.sheet(itemGid).tables[0]?.range?.startColumnIndex).toBe(
      expectedSheetLayout.startTableColIndex,
    );
    expect(result).toMatchObject({
      landed: [
        { subject: { table: "item", column: nameAt(2), path: "before" } },
      ],
    });
  });

  it("moves a column after the same run's column deletion, counting from where it left the columns", () => {
    const service = seedFixture();
    const remaining = columnNames.filter((name) => name !== "optionalNote");
    const moved = Val.assert(remaining.at(-1), "last remaining column");
    const target = Val.assert(remaining[1], "second remaining column");

    runItem({
      isSending: true,
      alterations: {
        tables: {
          item: {
            columns: {
              optionalNote: { deleted: true },
              [moved]: { before: target },
            },
          },
        },
      },
    });

    const expected = remaining.filter((name) => name !== moved);
    expected.splice(1, 0, moved);
    expect(headersOf(itemSheetRows(service))).toEqual(
      expected.map((name) => itemColumns[name].header),
    );
  });

  it("skips a before the live order already satisfies", () => {
    const service = seedFixture();
    const before = itemSheetRows(service);

    const { result } = runItem({
      isSending: true,
      alterations: moveAlterations(nameAt(0), nameAt(1)),
    });

    expect(itemSheetRows(service)).toEqual(before);
    expect(result).toMatchObject({ plan: [], landed: [] });
  });

  it("stops partway when Google refuses the move's batch over a taller Table beside it, the deletion before it landed", () => {
    const { colIdRowIndex } = expectedSheetLayout;
    let sentBatches = 0;
    const service = stubSheetsService({
      sheets: [itemSheet()],
      // Added after the fetch, so only Google sees it split.
      afterEachBatchUpdate(spreadsheet) {
        sentBatches += 1;
        if (sentBatches !== 1) return;
        const sheet = Val.assert(
          spreadsheet.sheets.find(({ sheetId }) => sheetId === itemGid),
          "Item sheet",
        );
        const item = Val.assert(sheet.tables[0], "Item Table");
        sheet.tables.push({
          tableId: "beside",
          name: "Beside",
          startRowIndex: colIdRowIndex,
          endRowIndex: item.endRowIndex + 4,
          startColumnIndex: item.endColumnIndex + 1,
          endColumnIndex: item.endColumnIndex + 3,
        });
      },
    });
    const remaining = columnNames.filter((name) => name !== "optionalNote");
    const moved = Val.assert(remaining.at(-1), "last remaining column");
    const target = Val.assert(remaining[1], "second remaining column");

    const { result, files } = runItem({
      isSending: true,
      alterations: {
        tables: {
          item: {
            columns: {
              optionalNote: { deleted: true },
              [moved]: { before: target },
            },
          },
        },
      },
    });

    expect(headersOf(itemSheetRows(service))).toEqual(
      remaining.map((name) => itemColumns[name].header),
    );
    expect(files).toBeUndefined();
    expect(result.outcome).toEqual({
      kind: "stoppedPartway",
      landedFlushes: 1,
      googleRefusal: expect.stringContaining(
        'Moving a column of Table "Item" on sheet "Item" would shift only part of Table "Beside". Move "Beside" so that it sits entirely within rows',
      ),
    });
  });

  describe("a moves group", () => {
    function groupAlterations(
      moves: { columns: ItemColumnName[]; before: ItemColumnName }[],
    ): PendingAlterations {
      return { tables: { item: { moves } } };
    }

    it("moves a two-column group, in its order, to sit before its target, and a rerun moves nothing", () => {
      const service = seedFixture();
      const [first, second, third, fourth, ...rest] = columnNames;
      const alterations = groupAlterations([
        {
          columns: [Val.assert(fourth, "fourth"), Val.assert(third, "third")],
          before: Val.assert(second, "second"),
        },
      ]);

      const { result, files } = runItem({ isSending: true, alterations });

      const moved = [first, fourth, third, second, ...rest].map(
        (name) => itemColumns[Val.assert(name, "item column")].header,
      );
      expect(headersOf(itemSheetRows(service))).toEqual(moved);
      expect(files?.alterations).toContain("makePendingAlterations({})");
      expect(result).toMatchObject({
        outcome: { kind: "sent" },
        plan: [{ subject: { table: "item", path: "moves[0]" } }],
        landed: [{ subject: { table: "item", path: "moves[0]" } }],
      });

      const rerun = runItem({ isSending: true, alterations });

      expect(headersOf(itemSheetRows(service))).toEqual(moved);
      expect(rerun.result).toMatchObject({
        plan: [],
        alreadyMet: [{ subject: { table: "item", path: "moves[0]" } }],
      });
    });

    it("applies groups in list order, each holding once all have moved", () => {
      const service = seedFixture();
      const [first, second, third, fourth, ...rest] = columnNames;

      runItem({
        isSending: true,
        alterations: groupAlterations([
          {
            columns: [Val.assert(third, "third")],
            before: Val.assert(second, "second"),
          },
          {
            columns: [Val.assert(first, "first")],
            before: Val.assert(fourth, "fourth"),
          },
        ]),
      });

      expect(headersOf(itemSheetRows(service))).toEqual(
        [third, second, first, fourth, ...rest].map(
          (name) => itemColumns[Val.assert(name, "item column")].header,
        ),
      );
    });

    function refusalsOf(alterations: PendingAlterations): unknown {
      const service = seedFixture();
      const { result, files } = runItem({ isSending: true, alterations });
      expect(files).toBeUndefined();
      expect(service.batchUpdateCount()).toBe(0);
      expect(result.outcome).toEqual({ kind: "refused" });
      return result.refusals;
    }

    it("refuses two groups with the same before, naming both", () => {
      expect(
        refusalsOf(
          groupAlterations([
            { columns: [nameAt(2)], before: nameAt(0) },
            { columns: [nameAt(3)], before: nameAt(0) },
          ]),
        ),
      ).toEqual([
        {
          kind: "clashingMoves",
          clash: "sameBefore",
          table: "item",
          path: "moves[1].before",
          otherPath: "moves[0].before",
          column: nameAt(0),
        },
      ]);
    });

    it("refuses a column in two groups, naming both", () => {
      expect(
        refusalsOf(
          groupAlterations([
            { columns: [nameAt(2)], before: nameAt(0) },
            { columns: [nameAt(3), nameAt(2)], before: nameAt(1) },
          ]),
        ),
      ).toEqual([
        {
          kind: "clashingMoves",
          clash: "repeatedColumn",
          table: "item",
          path: "moves[1].columns[1]",
          otherPath: "moves[0].columns[0]",
          column: nameAt(2),
        },
      ]);
    });

    it("refuses a group whose before is a column some group moves, its own included", () => {
      expect(
        refusalsOf(
          groupAlterations([
            { columns: [nameAt(2)], before: nameAt(0) },
            { columns: [nameAt(3)], before: nameAt(2) },
            { columns: [nameAt(1)], before: nameAt(1) },
          ]),
        ),
      ).toEqual([
        {
          kind: "clashingMoves",
          clash: "movedBefore",
          table: "item",
          path: "moves[1].before",
          otherPath: "moves[0].columns[0]",
          column: nameAt(2),
        },
        {
          kind: "clashingMoves",
          clash: "movedBefore",
          table: "item",
          path: "moves[2].before",
          otherPath: "moves[2].columns[0]",
          column: nameAt(1),
        },
      ]);
    });

    it("refuses a group's column that carries its own before, naming both", () => {
      expect(
        refusalsOf({
          tables: {
            item: {
              columns: { [nameAt(2)]: { before: nameAt(1) } },
              moves: [{ columns: [nameAt(2)], before: nameAt(0) }],
            },
          },
        }),
      ).toEqual([
        {
          kind: "clashingMoves",
          clash: "ownBefore",
          table: "item",
          path: "moves[0].columns[0]",
          otherPath: `${nameAt(2)}.before`,
          column: nameAt(2),
        },
      ]);
    });

    it("refuses a group naming a column the same run deletes", () => {
      expect(
        refusalsOf({
          tables: {
            item: {
              columns: { [nameAt(1)]: { deleted: true } },
              moves: [{ columns: [nameAt(3), nameAt(1)], before: nameAt(0) }],
            },
          },
        }),
      ).toMatchObject([
        {
          kind: "touchesDeletion",
          deleted: { table: "item", column: nameAt(1) },
          touchedBy: "tables.item.moves[0].columns[1]",
        },
      ]);
      expect(
        refusalsOf({
          tables: {
            item: {
              columns: { [nameAt(0)]: { deleted: true } },
              moves: [{ columns: [nameAt(3)], before: nameAt(0) }],
            },
          },
        }),
      ).toMatchObject([
        {
          kind: "touchesDeletion",
          deleted: { table: "item", column: nameAt(0) },
          touchedBy: "tables.item.moves[0].before",
        },
      ]);
    });
  });

  it("throws before sending anything on moving the last column to the front", () => {
    const service = seedFixture();

    expect(() =>
      runItem({
        isSending: true,
        alterations: moveAlterations(nameAt(columnNames.length - 1), nameAt(0)),
      }),
    ).toThrow(/its last column to the front/);
    expect(service.batchUpdateCount()).toBe(0);
  });
});
