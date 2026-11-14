import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { type PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  itemColumns,
  runItem,
  seedFixture,
} from "./syncCoordinatorTestSupport";

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
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
