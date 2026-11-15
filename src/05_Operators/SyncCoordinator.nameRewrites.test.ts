import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

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

function alterationsSource(body: string[]): string {
  return [
    "// sync rewrites this file: each alteration that lands is removed, and comments here won't survive a run.",
    'import { makePendingAlterations } from "@byronbroughten/sheets-framework";',
    "",
    "export const pendingAlterations = makePendingAlterations({",
    ...body,
    "});",
    "",
  ].join("\n");
}

describe("SyncCoordinator.run name rewrites", () => {
  const { optionalNote, requiredCount } = itemColumns;

  it("rewrites a column a person retitled, in its key and each before, on a bare run", () => {
    seedFixture({
      item: {
        columns: { requiredCount: { ...requiredCount, header: "Tally" } },
      },
    });

    const { result, files } = runItem({
      isSending: false,
      alterations: {
        tables: {
          item: {
            columns: { requiredCount: { emptyValueAllowed: true } },
            newColumns: [{ header: "Extra", before: "requiredCount" }],
          },
        },
      },
    });

    const emitted = Val.assert(files, "emitted files");
    expect(emitted.alterations).toBe(
      alterationsSource([
        '  "tables": {',
        '    "item": {',
        '      "columns": {',
        '        "tally": { "emptyValueAllowed": true }',
        "      },",
        '      "newColumns": [{"header":"Extra","before":"tally"}]',
        "    }",
        "  }",
      ]),
    );
    expect(result.configs).toMatchObject({
      renamed: [
        {
          table: "item",
          from: "requiredCount",
          to: "tally",
          headerFrom: requiredCount.header,
          headerTo: "Tally",
        },
      ],
      namesMoved: [],
      rewrites: [
        {
          from: "item.columns.requiredCount",
          to: "item.columns.tally",
          why: `${requiredCount.columnId}, which a person retitled "Tally"`,
        },
        {
          from: 'item.newColumns[0].before "requiredCount"',
          to: '"tally"',
        },
      ],
    });
    expect(result.plan).toContainEqual(
      expect.objectContaining({
        subject: {
          table: "item",
          column: "tally",
          rewrittenFrom: {
            wrote: "requiredCount",
            columnId: requiredCount.columnId,
          },
          path: "emptyValueAllowed",
        },
      }),
    );
  });

  // The conflict holds back the header the configs record, but the column's name still follows the live one.
  it("rewrites a Table and a column a person renamed on a stopped run", () => {
    seedFixture({
      item: {
        name: "Stock",
        columns: { optionalNote: { ...optionalNote, header: "Theirs" } },
      },
    });

    const { result, files } = runItem({
      isSending: true,
      alterations: {
        tables: { item: { columns: { optionalNote: { header: "Mine" } } } },
      },
    });

    expect(result.outcome).toEqual({ kind: "stopped" });
    expect(Val.assert(files, "emitted files").alterations).toBe(
      alterationsSource([
        '  "tables": {',
        '    "stock": {',
        '      "columns": {',
        '        "theirs": { "header": "Mine" }',
        "      }",
        "    }",
        "  }",
      ]),
    );
    expect(result.configs.rewrites).toEqual([
      {
        from: "item",
        to: "stock",
        why: "tableId item, which a person renamed",
      },
      {
        from: "item.columns.optionalNote",
        to: "stock.columns.theirs",
        why: `${optionalNote.columnId}, which a person retitled "Theirs"`,
      },
    ]);
    expect(result.stops).toMatchObject([
      {
        kind: "conflict",
        subject: {
          table: "stock",
          column: "theirs",
          rewrittenFrom: {
            wrote: "optionalNote",
            columnId: optionalNote.columnId,
          },
          path: "header",
        },
      },
    ]);
  });

  it("shows a pending entry under a renamed Table's new name when its column kept its name", () => {
    seedFixture({ item: { name: "Stock" } });

    const { result, files } = runItem({
      isSending: false,
      alterations: {
        tables: {
          item: { columns: { optionalNote: { emptyValueAllowed: false } } },
        },
      },
    });

    expect(Val.assert(files, "emitted files").alterations).toBe(
      alterationsSource([
        '  "tables": {',
        '    "stock": {',
        '      "columns": {',
        '        "optionalNote": { "emptyValueAllowed": false }',
        "      }",
        "    }",
        "  }",
      ]),
    );
    expect(result.plan).toContainEqual(
      expect.objectContaining({
        subject: {
          table: "stock",
          column: "optionalNote",
          path: "emptyValueAllowed",
        },
      }),
    );
  });

  it("follows each column's ID across a swap of two headers, and lists both names as moved", () => {
    seedFixture({
      item: {
        columns: {
          optionalNote: { ...optionalNote, header: requiredCount.header },
          requiredCount: { ...requiredCount, header: optionalNote.header },
        },
      },
    });

    const { result, files } = runItem({
      isSending: false,
      alterations: {
        tables: {
          item: { columns: { optionalNote: { emptyValueAllowed: false } } },
        },
      },
    });

    expect(Val.assert(files, "emitted files").alterations).toBe(
      alterationsSource([
        '  "tables": {',
        '    "item": {',
        '      "columns": {',
        '        "requiredCount": { "emptyValueAllowed": false }',
        "      }",
        "    }",
        "  }",
      ]),
    );
    expect(result.configs.namesMoved).toEqual([
      {
        table: "item",
        column: "optionalNote",
        wasId: optionalNote.columnId,
        nowId: requiredCount.columnId,
      },
      {
        table: "item",
        column: "requiredCount",
        wasId: requiredCount.columnId,
        nowId: optionalNote.columnId,
      },
    ]);
    expect(result.configs.rewrites).toEqual([
      {
        from: "item.columns.optionalNote",
        to: "item.columns.requiredCount",
        why: `${optionalNote.columnId}, which a person retitled "${requiredCount.header}"`,
      },
    ]);
  });

  it("leaves alterations.ts untouched when nothing it names was renamed", () => {
    seedFixture();

    const { files } = runItem({
      isSending: false,
      alterations: {
        tables: {
          item: { columns: { optionalNote: { emptyValueAllowed: false } } },
        },
      },
    });

    expect(Val.assert(files, "emitted files").alterations).toBeUndefined();
  });
});
