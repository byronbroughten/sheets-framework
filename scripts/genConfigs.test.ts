import { readFileSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

import { Str } from "@byronbroughten/utils/str";
import { beforeAll, describe, expect, it } from "vitest";

import { appConfigs } from "../dev/generated/appConfigs.ts";
import type { IdNameMapGeneric } from "../src/01_SpreadsheetSchema/configReaders/idNameMapTypes.ts";
import { columnIdsByName } from "../src/05_Operators/configFileSource.ts";
import { parseIdNameMapFileSource } from "../src/testSupport/parseIdNameMapFileSource.ts";
import {
  alterationsPath,
  alterationsText,
  appConfigsText,
  idNameMapPath,
} from "./genConfigs.ts";

const devGeneratedDir = fileURLToPath(
  new URL("../dev/generated", import.meta.url),
);
const devIdNameMapPath = idNameMapPath(devGeneratedDir);
const devAlterationsPath = alterationsPath(devGeneratedDir);

describe("the dev package's generated files", () => {
  const { tableConfigs, columnConfigs } = appConfigs;
  let idNameMap: IdNameMapGeneric;
  beforeAll(() => {
    idNameMap = parseIdNameMapFileSource(
      readFileSync(devIdNameMapPath, "utf8"),
    );
  });

  it("give each configured Table the same name, tab and column IDs in the ID-name map", () => {
    Object.entries(tableConfigs).forEach(([tableKey, tableConfig]) => {
      expect(idNameMap.tables[tableConfig.tableId]).toEqual({
        name: tableKey,
        sheetGid: tableConfig.sheetGid,
        columns: columnIdsByName(
          columnConfigs[tableKey as keyof typeof columnConfigs],
        ),
      });
    });
  });

  it("give columns in the ID-name map only to configured Tables", () => {
    const tableIdsWithColumns = Object.entries(idNameMap.tables).flatMap(
      ([tableId, table]) => (table.columns === undefined ? [] : [tableId]),
    );
    expect(tableIdsWithColumns.sort()).toEqual(
      Object.values(tableConfigs)
        .map((tableConfig) => tableConfig.tableId)
        .sort(),
    );
  });

  it("name each tab in the ID-name map by its title camelCased", () => {
    Object.values(idNameMap.tabs).forEach((tab) => {
      expect(tab.name).toBe(Str.sentenceToCamelCase(tab.title));
    });
  });

  it("place each Table in the ID-name map on a tab that lists it", () => {
    Object.entries(idNameMap.tables).forEach(([tableId, table]) => {
      expect(idNameMap.tabs[String(table.sheetGid)]?.tableIds).toContain(
        tableId,
      );
    });
  });
});

describe("appConfigsText", () => {
  const text = appConfigsText();

  it("builds appConfigs from the sibling config files", () => {
    expect(text).toContain('import { columnConfigs } from "./columnConfigs";');
    expect(text).toContain(
      'import { retiredIdPrefixes } from "./retiredIdPrefixes";',
    );
    expect(text).toContain('import { tabConfigs } from "./tabConfigs";');
    expect(text).toContain(
      'import { conditionalFormatConfigs } from "./conditionalFormatConfigs";',
    );
    expect(text).toContain('import { tableConfigs } from "./tableConfigs";');
    expect(text).toContain('import { valueConfigs } from "./valueConfigs";');
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

  it("registers the sibling ID-name map's type", () => {
    expect(text).toContain('import type { IdNameMap } from "./idNameMap";');
  });
});

describe("alterationsPath", () => {
  it("puts alterations.ts beside the generated directory", () => {
    expect(alterationsPath(join("app", "src", "generated"))).toBe(
      join("app", "src", "alterations.ts"),
    );
  });
});

describe("alterationsText", () => {
  it("holds an empty makePendingAlterations call from the framework's entry", () => {
    const text = alterationsText();
    expect(text).toContain(
      'import { makePendingAlterations } from "@byronbroughten/sheets-framework";',
    );
    expect(text).toContain(
      "export const pendingAlterations = makePendingAlterations({});",
    );
  });

  it("matches the dev package's own alterations.ts", () => {
    expect(readFileSync(devAlterationsPath, "utf8")).toBe(alterationsText());
  });
});
