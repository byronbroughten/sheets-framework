import { describe, expect, it } from "vitest";

import { assertType, type IsExactly } from "../../testSupport/typeAssertions";
import { installedConfigs } from "./configRegister";
import {
  configTableNames,
  getTableTraitByGid,
  getTableTraitByName,
  tableConfigsByName,
  tableKeysByGid,
  type TableNameSimple,
} from "./tableConfigsTypes";

const { tableConfigs } = installedConfigs();

describe("tableConfigs lookups", () => {
  it("names a Table by its tableConfigs key", () => {
    assertType<IsExactly<TableNameSimple, keyof typeof tableConfigs>>(true);
    expect(configTableNames()).toEqual(Object.keys(tableConfigs));
  });

  it("reads a trait by Table key, the key itself included", () => {
    expect(getTableTraitByName("item", "sheetGid")).toBe(
      tableConfigs.item.sheetGid,
    );
    expect(getTableTraitByName("item", "tableId")).toBe(
      tableConfigs.item.tableId,
    );
    expect(getTableTraitByName("item", "tableKey")).toBe("item");
  });

  it("reads a Table's config, key included, by its sheet GID", () => {
    const { sheetGid } = tableConfigs.log;
    expect(tableKeysByGid().get(sheetGid)).toEqual(["log"]);
    expect(getTableTraitByGid(sheetGid, "idPrefix")).toBe(
      tableConfigs.log.idPrefix,
    );
  });

  it("throws naming a GID no Table config holds", () => {
    expect(() => getTableTraitByGid(-1, "tableKey")).toThrow("gid -1");
  });

  it("reads an arbitrary name where the entry may be absent", () => {
    expect(tableConfigsByName()["item"]?.hasIdColumn).toBe(true);
    expect(tableConfigsByName()["nope"]).toBeUndefined();
  });
});
