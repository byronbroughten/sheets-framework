import { describe, expect, it } from "vitest";

import { columnConfigsByName } from "./columnConfigsTypes";
import { assertFloorMatchesSeed } from "./floorSeedCheck";
import type { ColumnConfigsGeneric, ColumnConfigStored } from "./makeConfigs";
import { tableConfigsByName } from "./tableConfigsTypes";

describe("assertFloorMatchesSeed", () => {
  const tableConfigs = tableConfigsByName();

  function floorColumnConfigs(): ColumnConfigsGeneric {
    return JSON.parse(JSON.stringify(columnConfigsByName()));
  }

  function floorColumn(
    configs: ColumnConfigsGeneric,
    tableName: string,
    columnName: string,
  ): ColumnConfigStored {
    const column = configs[tableName]?.[columnName];
    if (column === undefined) {
      throw new Error(`No generated column ${tableName}.${columnName}.`);
    }
    return column;
  }

  it("passes the generated floor entries against the floor seed", () => {
    expect(() =>
      assertFloorMatchesSeed(tableConfigs, floorColumnConfigs()),
    ).not.toThrow();
  });

  it("throws naming a floor tab with no entry", () => {
    const { valueConfig: _valueConfig, ...withoutValueConfig } = tableConfigs;

    expect(() =>
      assertFloorMatchesSeed(withoutValueConfig, floorColumnConfigs()),
    ).toThrow('Floor tab "valueConfig" has no floor entry');
  });

  it("throws naming the sheet, the column and both headers when a floor column's header differs from the seed's", () => {
    const configs = floorColumnConfigs();
    const header = floorColumn(configs, "columnConfig", "header");
    header.header = "Heading";

    expect(() => assertFloorMatchesSeed(tableConfigs, configs)).toThrow(
      `Floor column "Header" on "columnConfig" (column ID "${header.columnId}") has header "Heading" where the floor seed has "Header".`,
    );
  });

  it("throws when a floor column's valueName isn't the one the seed's column type implies", () => {
    const configs = floorColumnConfigs();
    const tableId = floorColumn(configs, "columnConfig", "tableId");
    tableId.valueName = "number";

    expect(() => assertFloorMatchesSeed(tableConfigs, configs)).toThrow(
      `Floor column "Table ID" on "columnConfig" (column ID "${tableId.columnId}") has valueName "number" where the floor seed's column type TEXT implies "string".`,
    );
  });

  it("throws when a BOOLEAN floor column's valueName isn't checkbox", () => {
    const configs = floorColumnConfigs();
    const letApiAccess = floorColumn(configs, "tableConfig", "letApiAccess");
    letApiAccess.valueName = "boolean";

    expect(() => assertFloorMatchesSeed(tableConfigs, configs)).toThrow(
      `Floor column "Let api access" on "tableConfig" (column ID "${letApiAccess.columnId}") has valueName "boolean" where the floor seed's column type BOOLEAN implies "checkbox".`,
    );
  });

  it("throws when a floor column's emptyValueAllowed differs from the seed's", () => {
    const configs = floorColumnConfigs();
    const tableMenuSpace = floorColumn(
      configs,
      "spreadsheetConfig",
      "tableMenuSpace",
    );
    tableMenuSpace.emptyValueAllowed = true;

    expect(() => assertFloorMatchesSeed(tableConfigs, configs)).toThrow(
      `Floor column "Table menu space" on "spreadsheetConfig" (column ID "${tableMenuSpace.columnId}") has emptyValueAllowed true where the floor seed has false.`,
    );
  });

  it("throws naming a seeded floor column with no entry", () => {
    const configs = floorColumnConfigs();
    delete configs.columnConfig?.emptyValueAllowed;

    expect(() => assertFloorMatchesSeed(tableConfigs, configs)).toThrow(
      'Floor column "Empty value allowed" on "columnConfig" has no floor entry.',
    );
  });

  it("passes a floor entry with a Custom default value and one without", () => {
    const configs = floorColumnConfigs();
    floorColumn(configs, "tableConfig", "sheetTitle").customDefaultValue =
      "Untitled";
    floorColumn(configs, "columnConfig", "tableName").customDefaultValue =
      null;

    expect(() => assertFloorMatchesSeed(tableConfigs, configs)).not.toThrow();
  });

  it("passes a live column the seed doesn't declare", () => {
    const configs = floorColumnConfigs();
    const columnConfigTab = configs.columnConfig ?? {};
    columnConfigTab.notes = {
      columnId: "c:ccf:notes01",
      header: "Notes",
      valueName: "string",
      isFormula: false,
      emptyValueAllowed: true,
      customDefaultValue: null,
    };

    expect(() => assertFloorMatchesSeed(tableConfigs, configs)).not.toThrow();
  });
});
