import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

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
