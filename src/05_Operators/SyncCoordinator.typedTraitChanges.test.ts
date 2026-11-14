import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  optionalNoteAlteration,
  runItem,
  seedFixture,
} from "./syncCoordinatorTestSupport";

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
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
