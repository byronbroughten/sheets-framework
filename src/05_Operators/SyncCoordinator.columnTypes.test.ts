import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { HeldFlushRefusal } from "../00_Source/RawSource/HeldFlushRefusal";
import type { DropdownStyleReset } from "../00_Source/RawSource/RawSource";
import type { ColumnTypeBaseline } from "../01_SpreadsheetSchema/columnBaseline";
import type { PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import type { FakeSheetsServiceOptions } from "../testSupport/fakeSheetsService";
import { SyncCoordinator } from "./SyncCoordinator";
import {
  itemColumns,
  noteColIndex,
  runItem,
  runWithNoAlterations,
  seedFixture,
} from "./syncCoordinatorTestSupport";

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
      isOtherADropdown?: boolean;
      service?: Omit<FakeSheetsServiceOptions, "sheets">;
    } = {},
  ): void {
    const dropdown = options.isOtherADropdown
      ? {
          columnTypes: { ...recordedTypes, [otherColIndex]: "DROPDOWN" },
          columnValidationConditionTypes: { [otherColIndex]: "ONE_OF_LIST" },
          columnValidationValues: { [otherColIndex]: listRule.values },
        }
      : { columnTypes: recordedTypes };
    seedFixture({ item: dropdown, service: options.service });
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
      isOtherADropdown: true,
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
      isOtherADropdown: true,
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
});
