import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { HeldFlushRefusal } from "../00_Source/RawSource/HeldFlushRefusal.ts";
import type { DropdownStyleReset } from "../00_Source/RawSource/RawSource.ts";
import type { ColumnTypeBaseline } from "../01_SpreadsheetSchema/columnBaseline.ts";
import type { PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState.ts";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals.ts";
import type { FakeSheetsServiceOptions } from "../testSupport/fakeSheetsService.ts";
import { SyncCoordinator } from "./SyncCoordinator.ts";
import {
  itemColumns,
  noteColIndex,
  runItem,
  runWithNoAlterations,
  seedFixture,
} from "./syncCoordinatorTestSupport.ts";

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("SyncCoordinator column types", () => {
  const columnNames = Object.keys(itemColumns);
  const countColIndex = columnNames.indexOf("requiredCount");
  // The one Item column besides ID, the count and the note.
  const otherColumn = Val.assert(
    columnNames.find(
      (name) => !["id", "requiredCount", "optionalNote"].includes(name),
    ),
    "Item's other column",
  );
  const otherColIndex = columnNames.indexOf(otherColumn);

  // Item's column types as the previous configs record them, so only an alteration changes one.
  const recordedTypes: Record<number, string> = {
    0: "TEXT",
    [countColIndex]: "DOUBLE",
    [otherColIndex]: "TEXT",
    [noteColIndex]: "TEXT",
  };
  const listRule = { condition: "ONE_OF_LIST", values: ["a", "b"] };

  function retype(columnType: ColumnTypeBaseline): PendingAlterations {
    return { tables: { item: { columns: { requiredCount: { columnType } } } } };
  }

  function seed(
    options: {
      dropdownColIndexes?: number[];
      service?: Omit<FakeSheetsServiceOptions, "sheets">;
    } = {},
  ): void {
    const dropdownColIndexes = options.dropdownColIndexes ?? [];
    function eachDropdown<VL>(value: VL): Record<number, VL> {
      return Object.fromEntries(
        dropdownColIndexes.map((colIndex) => [colIndex, value]),
      );
    }
    seedFixture({
      item: {
        columnTypes: { ...recordedTypes, ...eachDropdown("DROPDOWN") },
        columnValidationConditionTypes: eachDropdown("ONE_OF_LIST"),
        columnValidationValues: eachDropdown(listRule.values),
      },
      service: options.service,
    });
  }

  function liveColumnType(columnName: string): unknown {
    const coordinator = SyncCoordinator.init();
    runWithNoAlterations(coordinator);
    return coordinator.columnConfigOperator.traitReadings().live.tables.item
      ?.columns[columnName]?.columnType;
  }

  function send(
    columnType: ColumnTypeBaseline,
    dropdownStyleReset?: DropdownStyleReset,
  ) {
    return runItem({
      isSending: true,
      alterations: retype(columnType),
      dropdownStyleReset,
    }).result;
  }

  it("sends a type change, leaving the column with no rule", () => {
    seed();

    const result = send({ type: "CURRENCY", rule: null });

    expect(result.outcome.kind).toBe("sent");
    expect(result.plan).toMatchObject([
      { what: expect.stringContaining("CURRENCY") },
    ]);
    expect(result.plan[0]?.effects).toEqual([]);
    expect(liveColumnType("requiredCount")).toEqual({
      type: "CURRENCY",
      rule: null,
    });
  });

  it("sends a dropdown's rule with its type, listing the column it resets", () => {
    seed({ service: { dropdownStyleReset: "allow" } });

    const result = send({ type: "DROPDOWN", rule: listRule }, "allow");

    expect(result.outcome.kind).toBe("sent");
    expect(result.plan[0]?.effects).toEqual([
      { kind: "dropdownReset", columns: ["requiredCount"], setting: "allow" },
    ]);
    expect(liveColumnType("requiredCount")).toEqual({
      type: "DROPDOWN",
      rule: listRule,
    });
  });

  it('lists each dropdown the send resets and holds it under "ask"', () => {
    seed({
      dropdownColIndexes: [otherColIndex],
      service: {
        // The fake stands in for the Node host's gate, which holds the flush the planner listed.
        dropdownStyleReset: "allow",
        beforeEachBatchUpdate: () => {
          throw new HeldFlushRefusal({
            kind: "uncovered",
            landedFlushes: 0,
            flushCount: undefined,
          });
        },
      },
    });

    const result = send({ type: "CURRENCY", rule: null }, "ask");

    expect(result.plan[0]?.effects).toEqual([
      { kind: "dropdownReset", columns: [otherColumn], setting: "ask" },
    ]);
    expect(result.outcome.kind).toBe("stoppedBeforeSending");
  });

  it('sends a type change free under "allow", keeping the other dropdown', () => {
    seed({
      dropdownColIndexes: [otherColIndex],
      service: { dropdownStyleReset: "allow" },
    });

    const result = send({ type: "CURRENCY", rule: null }, "allow");

    expect(result.outcome.kind).toBe("sent");
    expect(result.plan[0]?.effects).toEqual([
      { kind: "dropdownReset", columns: [otherColumn], setting: "allow" },
    ]);
    expect(liveColumnType("requiredCount")).toEqual({
      type: "CURRENCY",
      rule: null,
    });
    expect(liveColumnType(otherColumn)).toEqual({
      type: "DROPDOWN",
      rule: listRule,
    });
  });

  // The count is a dropdown only on the sheet, so its retype is forced over it.
  function removeCountDropdown() {
    return runItem({
      isSending: true,
      alterations: retype({ type: "DOUBLE", rule: null }),
      forcedTraits: [
        { table: "item", column: "requiredCount", trait: "columnType" },
      ],
      dropdownStyleReset: "ask",
    }).result;
  }

  it('sends a dropdown\'s removal under "ask", listing no reset', () => {
    seed({
      dropdownColIndexes: [countColIndex],
      service: { dropdownStyleReset: "ask" },
    });

    const result = removeCountDropdown();

    expect(result.plan[0]?.effects).toEqual([]);
    expect(result.outcome.kind).toBe("sent");
    expect(liveColumnType("requiredCount")).toEqual({
      type: "DOUBLE",
      rule: null,
    });
  });

  it("lists no reset for two dropdowns one run removes", () => {
    seed({
      dropdownColIndexes: [countColIndex, otherColIndex],
      service: { dropdownStyleReset: "ask" },
    });
    const removed: ColumnTypeBaseline = { type: "TEXT", rule: null };

    const { result } = runItem({
      isSending: true,
      alterations: {
        tables: {
          item: {
            columns: {
              requiredCount: { columnType: removed },
              [otherColumn]: { columnType: removed },
            },
          },
        },
      },
      forcedTraits: [
        { table: "item", column: "requiredCount", trait: "columnType" },
        { table: "item", column: otherColumn, trait: "columnType" },
      ],
      dropdownStyleReset: "ask",
    });

    expect(result.plan.map(({ effects }) => effects)).toEqual([[], []]);
    expect(result.outcome.kind).toBe("sent");
  });

  it('holds a dropdown\'s removal under "ask" when another dropdown survives it', () => {
    seed({
      dropdownColIndexes: [countColIndex, otherColIndex],
      service: { dropdownStyleReset: "ask" },
    });

    const otherHeader = Val.assert(
      itemColumns[otherColumn as keyof typeof itemColumns],
      "the other column's config",
    ).header;

    expect(() => removeCountDropdown()).toThrow(
      `would reset the style and colours of its dropdowns ${otherHeader}.`,
    );
    expect(liveColumnType("requiredCount")).toEqual({
      type: "DROPDOWN",
      rule: listRule,
    });
  });
});
