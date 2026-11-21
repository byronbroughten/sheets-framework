import { Val } from "@byronbroughten/utils/val";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { stubLogger } from "../testSupport/fakeAppsScriptGlobals.ts";
import {
  itemColumns,
  runItem,
  seedFixture,
} from "./syncCoordinatorTestSupport.ts";

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
            newColumns: [
              { columns: [{ header: "Extra" }], before: "requiredCount" },
            ],
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
        '      "newColumns": [{"columns":[{"header":"Extra"}],"before":"tally"}]',
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

  it("rewrites retitled columns inside a pending moves group, in its columns and its before", () => {
    seedFixture({
      item: {
        columns: {
          requiredCount: { ...requiredCount, header: "Tally" },
          optionalNote: { ...optionalNote, header: "Remark" },
        },
      },
    });

    const { result, files } = runItem({
      isSending: false,
      alterations: {
        tables: {
          item: {
            moves: [{ columns: ["requiredCount"], before: "optionalNote" }],
          },
        },
      },
    });

    expect(Val.assert(files, "emitted files").alterations).toContain(
      '"moves": [{"columns":["tally"],"before":"remark"}]',
    );
    expect(result.configs?.rewrites).toEqual([
      { from: 'item.moves[0].columns[0] "requiredCount"', to: '"tally"' },
      { from: 'item.moves[0].before "optionalNote"', to: '"remark"' },
    ]);
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
        kind: "column",
        table: "item",
        column: "optionalNote",
        wasId: optionalNote.columnId,
        nowId: requiredCount.columnId,
      },
      {
        kind: "column",
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

  it("rewrites a renamed column's braces and Table[Header] references in a pending formula, leaving a string literal", () => {
    seedFixture({
      item: {
        columns: { requiredCount: { ...requiredCount, header: "Tally" } },
      },
    });
    const written = `=\${requiredCount}+SUM(item[${requiredCount.header}])+\${item[@requiredCount]}&"\${requiredCount}"`;
    const now = '=${tally}+SUM(Item[Tally])+${item[@tally]}&"${requiredCount}"';

    const { result, files } = runItem({
      isSending: false,
      alterations: {
        tables: { item: { columns: { optionalNote: { formula: written } } } },
      },
    });

    expect(Val.assert(files, "emitted files").alterations).toContain(
      JSON.stringify(now),
    );
    expect(result.configs.rewrites).toEqual([
      {
        from: `item.columns.optionalNote.formula "${written}"`,
        to: `"${now}"`,
      },
    ]);
  });

  it("rewrites a renamed Table's braces and Table[Header] references in pending formulas, a new Table's included, keeping a bare brace", () => {
    seedFixture({ item: { name: "Stock" } });
    const written = `=\${requiredCount}+SUM(item[${requiredCount.header}])+\${item[@requiredCount]}`;
    const now = `=\${requiredCount}+SUM(Stock[${requiredCount.header}])+\${stock[@requiredCount]}`;

    const { result, files } = runItem({
      isSending: false,
      alterations: {
        tables: { item: { columns: { optionalNote: { formula: written } } } },
        newTabs: [
          {
            title: "Ledger",
            tables: [
              {
                name: "Payments",
                columns: [
                  { header: "Due", formula: "=SUM(${item[requiredCount]})" },
                ],
              },
            ],
          },
        ],
      },
    });

    const emitted = Val.assert(files, "emitted files").alterations;
    expect(emitted).toContain(JSON.stringify(now));
    expect(emitted).toContain(JSON.stringify("=SUM(${stock[requiredCount]})"));
    expect(result.configs.rewrites).toEqual([
      {
        from: "item",
        to: "stock",
        why: "tableId item, which a person renamed",
      },
      {
        from: `item.columns.optionalNote.formula "${written}"`,
        to: `"${now}"`,
      },
      {
        from: 'newTabs[0].tables[0].columns[0].formula "=SUM(${item[requiredCount]})"',
        to: '"=SUM(${stock[requiredCount]})"',
      },
    ]);
  });

  it("rewrites a renamed column in a pending rule's columns, condition values and handles", () => {
    seedFixture({
      item: {
        columns: { requiredCount: { ...requiredCount, header: "Tally" } },
      },
    });
    const rule = {
      columns: ["requiredCount" as const],
      condition: {
        type: "CUSTOM_FORMULA" as const,
        values: ["=${@requiredCount}>0"],
      },
      format: { backgroundColor: "#00ff00" },
    };
    const handle = "requiredCount | =${@requiredCount}>0 | background #ff0000";
    const handleNow = "tally | =${@tally}>0 | background #ff0000";
    // Its own handle: one both removed and replaced is refused before the read.
    const removed = "requiredCount | =${@requiredCount}>1 | background #ff0000";
    const removedNow = "tally | =${@tally}>1 | background #ff0000";

    const { result } = runItem({
      isSending: false,
      alterations: {
        tables: {
          item: {
            conditionalFormats: {
              add: [rule],
              remove: [removed],
              replace: { [handle]: rule },
            },
          },
        },
      },
    });

    expect(result.configs.rewrites).toEqual([
      {
        from: 'item.conditionalFormats.add[0].columns "requiredCount"',
        to: '"tally"',
      },
      {
        from: 'item.conditionalFormats.add[0].condition.values[0] "=${@requiredCount}>0"',
        to: '"=${@tally}>0"',
      },
      {
        from: `item.conditionalFormats.remove[0] "${removed}"`,
        to: `"${removedNow}"`,
      },
      {
        from: `item.conditionalFormats.replace "${handle}"`,
        to: `"${handleNow}"`,
      },
      {
        from: `item.conditionalFormats.replace["${handle}"].columns "requiredCount"`,
        to: '"tally"',
      },
      {
        from: `item.conditionalFormats.replace["${handle}"].condition.values[0] "=\${@requiredCount}>0"`,
        to: '"=${@tally}>0"',
      },
    ]);
  });

  it("rewrites a handle whose condition names a renamed Table", () => {
    seedFixture({ item: { name: "Stock" } });
    const handle = `optionalNote | =COUNT(item[${requiredCount.header}])>\${item[requiredCount]} | background #ff0000`;
    const handleNow = `optionalNote | =COUNT(Stock[${requiredCount.header}])>\${stock[requiredCount]} | background #ff0000`;

    const { result } = runItem({
      isSending: false,
      alterations: {
        tables: { item: { conditionalFormats: { remove: [handle] } } },
      },
    });

    expect(result.configs.rewrites).toContainEqual({
      from: `item.conditionalFormats.remove[0] "${handle}"`,
      to: `"${handleNow}"`,
    });
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
