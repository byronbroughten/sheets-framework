import { mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

import { appConfigs } from "../dev/generated/appConfigs.ts";
import { alterationsFileSource } from "../src/05_Operators/configFileSource.ts";
import { GeneratedDir } from "./GeneratedDir.ts";

const devGeneratedDir = new GeneratedDir(
  fileURLToPath(new URL("../dev/generated", import.meta.url)),
);
const missingFilesDir = mkdtempSync(join(tmpdir(), "generated-dir-"));
const fixtureAlterationsDir = mkdtempSync(join(tmpdir(), "generated-dir-"));

describe("GeneratedDir's paths", () => {
  const { path } = new GeneratedDir(join("app", "src", "generated"));

  it("put alterations.ts beside the generated directory", () => {
    expect(path.alterations).toBe(join("app", "src", "alterations.ts"));
  });

  it("put the ID-name map, appConfigs and each config file inside it", () => {
    expect(path.idNameMap).toBe(
      join("app", "src", "generated", "idNameMap.ts"),
    );
    expect(path.appConfigs).toBe(
      join("app", "src", "generated", "appConfigs.ts"),
    );
    expect(path.columnConfigs).toBe(
      join("app", "src", "generated", "columnConfigs.ts"),
    );
  });
});

describe("loadIntendedState", () => {
  it("reads a missing alterations.ts as an empty one", async () => {
    const generatedDir = new GeneratedDir(join(missingFilesDir, "generated"));
    expect(await generatedDir.loadIntendedState()).toEqual({ alterations: {} });
  });

  it("takes the pending alterations alterations.ts exports", async () => {
    const generatedDir = new GeneratedDir(
      join(fixtureAlterationsDir, "generated"),
    );
    writeFileSync(
      generatedDir.path.alterations,
      "export const pendingAlterations = { tables: { item: { delete: true } } };\n",
    );
    expect(await generatedDir.loadIntendedState()).toEqual({
      alterations: { tables: { item: { delete: true } } },
    });
  });
});

describe("loadPreviousIdNameMap", () => {
  it("reads a missing ID-name map as an empty map", () => {
    expect(new GeneratedDir(missingFilesDir).loadPreviousIdNameMap()).toEqual({
      tables: {},
      tabs: {},
    });
  });

  it("reads the ID-name map back from its source", () => {
    const { tables } = devGeneratedDir.loadPreviousIdNameMap();
    for (const [tableKey, { tableId }] of Object.entries(
      appConfigs.tableConfigs,
    )) {
      expect(tables[tableId]?.name).toBe(tableKey);
    }
  });
});

describe("appConfigsText", () => {
  const text = GeneratedDir.appConfigsText;

  it("builds appConfigs from the sibling config files", () => {
    expect(text).toContain(
      'import { columnConfigs } from "./columnConfigs.ts";',
    );
    expect(text).toContain(
      'import { retiredIdPrefixes } from "./retiredIdPrefixes.ts";',
    );
    expect(text).toContain('import { tabConfigs } from "./tabConfigs.ts";');
    expect(text).toContain(
      'import { conditionalFormatConfigs } from "./conditionalFormatConfigs.ts";',
    );
    expect(text).toContain('import { tableConfigs } from "./tableConfigs.ts";');
    expect(text).toContain('import { valueConfigs } from "./valueConfigs.ts";');
    expect(text).toContain(
      "export const appConfigs = {\n  tableConfigs,\n  tabConfigs,\n  conditionalFormatConfigs,\n  retiredIdPrefixes,\n  columnConfigs,\n  valueConfigs,\n};",
    );
  });

  it("brings the framework's entry into the program with a type-only import", () => {
    expect(text).toContain(
      'import type {} from "@byronbroughten/sheets-framework";',
    );
  });

  it("augments the package's Register with the appConfigs type", () => {
    expect(text).toMatch(
      /declare module "@byronbroughten\/sheets-framework" \{\s*interface Register \{\s*configs: typeof appConfigs;\s*idNameMap: IdNameMap;\s*\}\s*\}/,
    );
  });

  it("is what the framework's dev/generated/appConfigs.ts holds", () => {
    expect(
      readFileSync(
        fileURLToPath(
          new URL("../dev/generated/appConfigs.ts", import.meta.url),
        ),
        "utf8",
      ),
    ).toBe(text);
  });

  it("registers the sibling ID-name map's type", () => {
    expect(text).toContain('import type { IdNameMap } from "./idNameMap.ts";');
  });
});

describe("alterationsText", () => {
  it("is the empty alterations file source", () => {
    expect(GeneratedDir.alterationsText).toBe(alterationsFileSource({}));
  });

  it("matches the dev package's own alterations.ts", () => {
    expect(readFileSync(devGeneratedDir.path.alterations, "utf8")).toBe(
      GeneratedDir.alterationsText,
    );
  });
});
