import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { type PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState";
import { expectedSheetLayout } from "../testSupport/expectedSheetLayout";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  type FakeCellValue,
  type FakeSheetsService,
} from "../testSupport/fakeSheetsService";
import { parseIdNameMapFileSource } from "./parseIdNameMapFileSource";
import {
  emittedItemColumn,
  emptyValueAllowedOf,
  type ItemColumnName,
  itemColumns,
  itemGid,
  itemRowId,
  runItem,
  seedFixture,
} from "./syncCoordinatorTestSupport";

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
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

  it("creates a group's columns together, in order, before the column it names", () => {
    const service = seedFixture();
    const [, secondColumn] = Object.keys(itemColumns) as ItemColumnName[];
    const before = itemSheetRows(service);
    const endBefore = itemTableEnd(service);

    const { result, files } = runItem({
      isSending: true,
      alterations: {
        tables: {
          item: {
            newColumns: [
              {
                columns: [
                  { header: "Ref" },
                  { header: "Late fee", emptyValueAllowed: true },
                ],
                before: Val.assert(secondColumn, "second column"),
              },
            ],
          },
        },
      },
    });

    const emitted = Val.assert(files, "emitted files");
    const [firstHeader, ...restHeaders] = headersOf(before);
    expect(headersOf(itemSheetRows(service))).toEqual([
      firstHeader,
      "Ref",
      "Late fee",
      ...restHeaders,
    ]);
    expect(itemTableEnd(service)).toBe(endBefore + 2);
    expect(emptyValueAllowedOf(emitted, "ref")).toBe(false);
    expect(emptyValueAllowedOf(emitted, "lateFee")).toBe(true);
    expect(emitted.alterations).toContain("makePendingAlterations({})");
    expect(result).toMatchObject({
      outcome: { kind: "sent" },
      landed: [
        { kind: "createdColumn", column: "ref" },
        { kind: "createdColumn", column: "lateFee" },
      ],
    });
  });

  it("appends a group's columns in order at the Table's end", () => {
    const service = seedFixture();
    const before = itemSheetRows(service);

    runItem({
      isSending: true,
      alterations: {
        tables: {
          item: {
            newColumns: [
              { columns: [{ header: "Late fee" }, { header: "Ref" }] },
            ],
          },
        },
      },
    });

    expect(headersOf(itemSheetRows(service))).toEqual([
      ...headersOf(before),
      "Late fee",
      "Ref",
    ]);
  });

  describe("refusals naming a group's paths", () => {
    const [, secondColumn, thirdColumn] = Object.keys(
      itemColumns,
    ) as ItemColumnName[];
    const secondName = Val.assert(secondColumn, "second column");
    const thirdName = Val.assert(thirdColumn, "third column");

    function refusalsOf(alterations: PendingAlterations): unknown {
      const service = seedFixture();
      const { result, files } = runItem({ isSending: true, alterations });
      expect(files).toBeUndefined();
      expect(service.batchUpdateCount()).toBe(0);
      expect(result.outcome).toEqual({ kind: "refused" });
      return result.refusals;
    }

    it("refuses two items with the same before, naming both", () => {
      expect(
        refusalsOf({
          tables: {
            item: {
              newColumns: [
                { header: "Ref", before: secondName },
                { columns: [{ header: "Late fee" }], before: secondName },
              ],
            },
          },
        }),
      ).toEqual([
        {
          kind: "clashingPlacements",
          clash: "sameNewColumnsBefore",
          table: "item",
          path: "newColumns[1].before",
          otherPath: "newColumns[0].before",
          column: secondName,
        },
      ]);
    });

    it("refuses a before a moves group also puts its group before, naming both", () => {
      expect(
        refusalsOf({
          tables: {
            item: {
              newColumns: [
                { columns: [{ header: "Ref" }], before: secondName },
              ],
              moves: [{ columns: [thirdName], before: secondName }],
            },
          },
        }),
      ).toEqual([
        {
          kind: "clashingPlacements",
          clash: "newColumnsBeforeMoveTarget",
          table: "item",
          path: "newColumns[0].before",
          otherPath: "moves[0].before",
          column: secondName,
        },
      ]);
    });

    it("refuses a before a moves group moves, naming both", () => {
      expect(
        refusalsOf({
          tables: {
            item: {
              newColumns: [{ columns: [{ header: "Ref" }], before: thirdName }],
              moves: [{ columns: [thirdName], before: secondName }],
            },
          },
        }),
      ).toEqual([
        {
          kind: "clashingPlacements",
          clash: "newColumnsBeforeMovedColumn",
          table: "item",
          path: "newColumns[0].before",
          otherPath: "moves[0].columns[0]",
          column: thirdName,
        },
      ]);
    });

    it("refuses two headers giving one name, one in a group, by the group column's path", () => {
      expect(
        refusalsOf({
          tables: {
            item: {
              newColumns: [
                { header: "Late fee" },
                { columns: [{ header: "Ref" }, { header: "Late Fee" }] },
              ],
            },
          },
        }),
      ).toEqual([
        {
          kind: "repeatedNewColumn",
          table: "item",
          path: "newColumns[1].columns[1].header",
          firstPath: "newColumns[0].header",
          column: "lateFee",
        },
      ]);
    });

    it("refuses a group column's trait not sendable yet by its path", () => {
      expect(
        refusalsOf({
          tables: {
            item: {
              newColumns: [
                {
                  columns: [{ header: "Ref" }, { header: "Fee", formula: "1" }],
                },
              ],
            },
          },
        }),
      ).toEqual([
        {
          kind: "notSendableYet",
          table: "item",
          path: "newColumns[0].columns[1].formula",
        },
      ]);
    });
  });

  it("stops on a group placed before the Table's first column, naming the group's before", () => {
    const service = seedFixture();
    const [firstColumn] = Object.keys(itemColumns) as ItemColumnName[];
    const firstName = Val.assert(firstColumn, "first column");

    const { result } = runItem({
      isSending: true,
      alterations: {
        tables: {
          item: {
            newColumns: [
              {
                columns: [{ header: "Ref" }, { header: "Late fee" }],
                before: firstName,
              },
            ],
          },
        },
      },
    });

    expect(service.batchUpdateCount()).toBe(0);
    expect(result).toMatchObject({
      outcome: { kind: "stopped" },
      stops: [
        {
          kind: "beforeFirstColumn",
          path: "newColumns[0].before",
          header: "Ref",
          before: firstName,
        },
        {
          kind: "beforeFirstColumn",
          path: "newColumns[0].before",
          header: "Late fee",
          before: firstName,
        },
      ],
    });
  });

  it("stops on a group column whose header the Table already has", () => {
    const service = seedFixture();
    const { header, columnId } = itemColumns.optionalNote;

    const { result } = runItem({
      isSending: true,
      alterations: {
        tables: {
          item: { newColumns: [{ columns: [{ header: "Ref" }, { header }] }] },
        },
      },
    });

    expect(service.batchUpdateCount()).toBe(0);
    expect(result).toMatchObject({
      outcome: { kind: "stopped" },
      stops: [
        { kind: "headerTaken", header, column: "optionalNote", columnId },
      ],
    });
  });

  describe("an entry placed before the Table's first column", () => {
    const [firstColumn] = Object.keys(itemColumns) as ItemColumnName[];
    const firstName = Val.assert(firstColumn, "first column");
    const beforeFirst: PendingAlterations = {
      tables: {
        item: {
          newColumns: [
            { header: "Late fee" },
            { header: "Ref", before: firstName },
          ],
        },
      },
    };
    const stop = {
      kind: "beforeFirstColumn",
      table: "item",
      path: "newColumns[1].before",
      header: "Ref",
      before: firstName,
    };

    it("stops a dry run, which reports it", () => {
      const service = seedFixture();

      const { result } = runItem({
        isSending: false,
        alterations: beforeFirst,
      });

      expect(service.batchUpdateCount()).toBe(0);
      expect(result).toMatchObject({
        outcome: { kind: "stopped" },
        stops: [stop],
      });
    });

    it("stops a send before sending anything", () => {
      const service = seedFixture();

      const { result, files } = runItem({
        isSending: true,
        alterations: beforeFirst,
      });

      expect(service.batchUpdateCount()).toBe(0);
      expect(files?.alterations).toBeUndefined();
      expect(result).toMatchObject({
        outcome: { kind: "stopped" },
        stops: [stop],
      });
    });
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
