import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { IdPrefixesByTableId } from "../01_SpreadsheetSchema/idPrefixes";
import { expectedSheetLayout } from "../testSupport/expectedSheetLayout";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  type FakeCell,
  type FakeSheetsService,
} from "../testSupport/fakeSheetsService";
import {
  emittedItemColumn,
  itemColumns,
  itemGid,
  itemRowId,
  noteColIndex,
  runItem,
  seedFixture,
} from "./syncCoordinatorTestSupport";
import { TableConfigsOperator } from "./TableConfigsOperator";

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("SyncCoordinator.run Fill Row IDs", () => {
  const itemColumnNames = Object.keys(itemColumns);
  const idColIndex = itemColumnNames.indexOf("id");

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
