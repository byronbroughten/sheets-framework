import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import type { PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import { SyncCoordinator } from "./SyncCoordinator";
import {
  frame,
  itemColumns,
  itemGid,
  previousItemMap,
  runItem,
  seedFixture,
} from "./syncCoordinatorTestSupport";

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

const emptyAlterationsFile = [
  "// sync rewrites this file: each alteration that lands is removed, and comments here won't survive a run.",
  'import { makePendingAlterations } from "@byronbroughten/sheets-framework";',
  "",
  "export const pendingAlterations = makePendingAlterations({});",
  "",
].join("\n");

// The configs and the live sheet both already allow optionalNote empty.
const metNote = {
  subject: { table: "item", column: "optionalNote", path: "emptyValueAllowed" },
  what: "already true",
};

const metNoteAlteration: PendingAlterations = {
  tables: {
    item: { columns: { optionalNote: { emptyValueAllowed: true } } },
  },
};

// optionalNote already met, requiredCount still to send.
const metNoteAndCountAlterations: PendingAlterations = {
  tables: {
    item: {
      columns: {
        optionalNote: { emptyValueAllowed: true },
        requiredCount: { emptyValueAllowed: true },
      },
    },
  },
};

describe("SyncCoordinator already-met alterations", () => {
  it("removes an alteration already met on both sides when nothing else is sent", () => {
    seedFixture();
    expect(itemColumns.optionalNote.emptyValueAllowed).toBe(true);

    const { result, files } = runItem({
      isSending: true,
      alterations: metNoteAlteration,
    });

    expect(Val.assert(files, "emitted files").alterations).toBe(
      emptyAlterationsFile,
    );
    expect(result).toMatchObject({
      mode: "send",
      outcome: {
        kind: "sent",
        summary: "nothing to send; 1 alteration already met",
      },
      plan: [],
      landed: [],
      alreadyMet: [metNote],
      isAlterationsEmpty: true,
    });
  });

  it("removes a met alteration beside a sent one, listing each apart", () => {
    seedFixture();
    expect(itemColumns.requiredCount.emptyValueAllowed).toBe(false);

    const { result, files } = runItem({
      isSending: true,
      alterations: metNoteAndCountAlterations,
    });

    expect(Val.assert(files, "emitted files").alterations).toBe(
      emptyAlterationsFile,
    );
    expect(result).toMatchObject({
      mode: "send",
      outcome: {
        kind: "sent",
        summary: "1 alteration landed; 1 alteration already met",
      },
      landed: [
        {
          kind: "send",
          subject: {
            table: "item",
            column: "requiredCount",
            path: "emptyValueAllowed",
          },
          what: "set to true",
        },
      ],
      alreadyMet: [metNote],
      isAlterationsEmpty: true,
    });
  });

  it("leaves alterations.ts holding an entry only the sheet holds, beside a landed one", () => {
    seedFixture({
      item: {
        columns: {
          optionalNote: { ...itemColumns.optionalNote, header: "Remark" },
        },
      },
    });

    const { result, files } = runItem({
      isSending: true,
      alterations: {
        tables: {
          item: {
            columns: {
              optionalNote: { header: "Remark" },
              requiredCount: { emptyValueAllowed: true },
            },
          },
        },
      },
    });

    const written = Val.assert(files, "emitted files").alterations;
    expect(written).toContain('"header": "Remark"');
    expect(written).not.toContain("requiredCount");
    expect(result).toMatchObject({
      mode: "send",
      outcome: { kind: "sent", summary: "1 alteration landed" },
      alreadyMet: [],
      notSent: [],
      didntLand: [],
      isAlterationsEmpty: false,
    });
  });

  it("leaves alterations.ts holding an entry on a tab a person deleted, beside a landed one", () => {
    seedFixture();
    const previous = previousItemMap();
    const logGid = itemGid + 1;

    const { result, files } = SyncCoordinator.init().run({
      isSending: true,
      intendedState: {
        alterations: {
          tabs: { log: { title: "Logbook" } },
          tables: {
            item: { columns: { requiredCount: { emptyValueAllowed: true } } },
          },
        },
      },
      previousIdNameMap: {
        tables: previous.tables,
        tabs: {
          ...previous.tabs,
          [String(logGid)]: { name: "log", title: "Log", tableIds: [] },
        },
      },
      forcedTraits: [],
      frame,
    });

    const written = Val.assert(files, "emitted files").alterations;
    expect(written).toContain('"title": "Logbook"');
    expect(written).not.toContain("requiredCount");
    expect(result).toMatchObject({
      mode: "send",
      outcome: { kind: "sent", summary: "1 alteration landed" },
      alreadyMet: [],
      notSent: [],
      didntLand: [],
      isAlterationsEmpty: false,
    });
  });

  it("removes nothing on a dry run, yet lists the met alteration", () => {
    seedFixture();

    const { result, files } = runItem({
      isSending: false,
      alterations: metNoteAlteration,
    });

    expect(Val.assert(files, "emitted files").alterations).toBeUndefined();
    expect(result).toMatchObject({
      mode: "dry run",
      outcome: { kind: "ready" },
      alreadyMet: [metNote],
    });
  });

  it("removes nothing on a stopped send, yet lists the met alteration", () => {
    seedFixture({
      item: { columns: { requiredCount: { columnId: "", header: "Scratch" } } },
    });

    const { result, files } = runItem({
      isSending: true,
      alterations: metNoteAndCountAlterations,
    });

    expect(Val.assert(files, "emitted files").alterations).toBeUndefined();
    expect(result).toMatchObject({
      mode: "send",
      outcome: { kind: "stopped" },
      landed: [],
      alreadyMet: [metNote],
    });
  });
});
