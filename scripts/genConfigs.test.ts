import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

import { beforeAll, describe, expect, it } from "vitest";

import { appConfigs } from "../dev/generated/appConfigs.ts";
import {
  columnIdsByName,
  type IdNameMapGeneric,
} from "../src/05_Operators/configFileSource.ts";
import { parseIdNameMapFileSource } from "../src/testSupport/parseIdNameMapFileSource.ts";
import { appConfigsText, idNameMapPath } from "./genConfigs.ts";

const devIdNameMapPath = idNameMapPath(
  fileURLToPath(new URL("../dev/generated", import.meta.url)),
);

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

  it("builds appConfigs from the three sibling config files", () => {
    expect(text).toContain('import { columnConfigs } from "./columnConfigs";');
    expect(text).toContain(
      'import { retiredIdPrefixes, tabConfigs, tableConfigs } from "./tableConfigs";',
    );
    expect(text).toContain('import { valueConfigs } from "./valueConfigs";');
    expect(text).toContain(
      "export const appConfigs = {\n  tableConfigs,\n  tabConfigs,\n  retiredIdPrefixes,\n  columnConfigs,\n  valueConfigs,\n};",
    );
  });

  it("brings the framework's entry into the program with a type-only import", () => {
    expect(text).toContain(
      'import type {} from "@byronbroughten/sheets-framework";',
    );
  });

  it("augments the package's Register with the appConfigs type", () => {
    expect(text).toMatch(
      /declare module "@byronbroughten\/sheets-framework" \{\s*interface Register \{\s*configs: typeof appConfigs;\s*\}\s*\}/,
    );
  });
});
