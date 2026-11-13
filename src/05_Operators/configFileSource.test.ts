import { describe, expect, it } from "vitest";

import { alterationsFileSource } from "./configFileSource";

const header =
  "// sync rewrites this file: each alteration that lands is removed, and comments here won't survive a run.";
const importLine =
  'import { makePendingAlterations } from "@byronbroughten/sheets-framework";';

describe("alterationsFileSource", () => {
  it("writes empty alterations", () => {
    expect(alterationsFileSource({})).toBe(
      [
        header,
        importLine,
        ``,
        `export const pendingAlterations = makePendingAlterations({});`,
        ``,
      ].join("\n"),
    );
  });

  it("writes one line per leaf, opening a block for each nesting", () => {
    expect(
      alterationsFileSource({
        tables: {
          item: {
            name: "thing",
            columns: {
              note: { emptyValueAllowed: false },
              count: { header: "Count", formula: undefined },
            },
          },
          unit: { deleted: true },
        },
      }),
    ).toBe(
      [
        header,
        importLine,
        ``,
        `export const pendingAlterations = makePendingAlterations({`,
        `  "tables": {`,
        `    "item": {`,
        `      "name": "thing",`,
        `      "columns": {`,
        `        "note": { "emptyValueAllowed": false },`,
        `        "count": { "header": "Count" }`,
        `      }`,
        `    },`,
        `    "unit": { "deleted": true }`,
        `  }`,
        `});`,
        ``,
      ].join("\n"),
    );
  });
});
