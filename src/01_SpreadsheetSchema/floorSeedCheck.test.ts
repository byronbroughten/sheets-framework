import { describe, expect, it } from "vitest";

import { columnConfigsByName } from "./columnConfigsTypes";
import { assertFloorMatchesSeed } from "./floorSeedCheck";
import type { ColumnConfigsGeneric, ColumnConfigStored } from "./makeConfigs";
import { sheetConfigsByName } from "./sheetConfigsTypes";

describe("assertFloorMatchesSeed", () => {
  const sheetConfigs = sheetConfigsByName();

  function floorColumnConfigs(): ColumnConfigsGeneric {
    return JSON.parse(JSON.stringify(columnConfigsByName()));
  }

  function floorColumn(
    configs: ColumnConfigsGeneric,
    sheetName: string,
    columnName: string,
  ): ColumnConfigStored {
    const column = configs[sheetName]?.[columnName];
    if (column === undefined) {
      throw new Error(`No generated column ${sheetName}.${columnName}.`);
    }
    return column;
  }

  it("passes the generated floor entries against the floor seed", () => {
    expect(() =>
      assertFloorMatchesSeed(sheetConfigs, floorColumnConfigs()),
    ).not.toThrow();
  });

  it("throws naming a floor tab with no entry", () => {
    const { valueConfig: _valueConfig, ...withoutValueConfig } = sheetConfigs;

    expect(() =>
      assertFloorMatchesSeed(withoutValueConfig, floorColumnConfigs()),
    ).toThrow('Floor tab "valueConfig" has no floor entry');
  });

  it("throws naming the sheet, the column and both headers when a floor column's header differs from the seed's", () => {
    const configs = floorColumnConfigs();
    const header = floorColumn(configs, "columnConfig", "header");
    header.header = "Heading";

    expect(() => assertFloorMatchesSeed(sheetConfigs, configs)).toThrow(
      `Floor column "Header" on "columnConfig" (column ID "${header.columnId}") has header "Heading" where the floor seed has "Header".`,
    );
  });

  it("throws when a floor column's valueName isn't the one the seed's column type implies", () => {
    const configs = floorColumnConfigs();
    const sheetGid = floorColumn(configs, "sheetConfig", "sheetGid");
    sheetGid.valueName = "string";

    expect(() => assertFloorMatchesSeed(sheetConfigs, configs)).toThrow(
      `Floor column "Sheet GID" on "sheetConfig" (column ID "${sheetGid.columnId}") has valueName "string" where the floor seed's column type DOUBLE implies "number".`,
    );
  });

  it("throws when a BOOLEAN floor column's valueName isn't checkbox", () => {
    const configs = floorColumnConfigs();
    const letApiAccess = floorColumn(configs, "sheetConfig", "letApiAccess");
    letApiAccess.valueName = "boolean";

    expect(() => assertFloorMatchesSeed(sheetConfigs, configs)).toThrow(
      `Floor column "Let api access" on "sheetConfig" (column ID "${letApiAccess.columnId}") has valueName "boolean" where the floor seed's column type BOOLEAN implies "checkbox".`,
    );
  });

  it("throws when a floor column's emptyValueAllowed differs from the seed's", () => {
    const configs = floorColumnConfigs();
    const idHeader = floorColumn(configs, "spreadsheetConfig", "idHeader");
    idHeader.emptyValueAllowed = true;

    expect(() => assertFloorMatchesSeed(sheetConfigs, configs)).toThrow(
      `Floor column "ID header" on "spreadsheetConfig" (column ID "${idHeader.columnId}") has emptyValueAllowed true where the floor seed has false.`,
    );
  });

  it("throws naming a seeded floor column with no entry", () => {
    const configs = floorColumnConfigs();
    delete configs.columnConfig?.emptyValueAllowed;

    expect(() => assertFloorMatchesSeed(sheetConfigs, configs)).toThrow(
      'Floor column "Empty value allowed" on "columnConfig" has no floor entry.',
    );
  });

  it("passes a floor entry with a Custom default value and one without", () => {
    const configs = floorColumnConfigs();
    floorColumn(configs, "sheetConfig", "sheetTitle").customDefaultValue =
      "Untitled";
    floorColumn(configs, "columnConfig", "sheetTitle").customDefaultValue =
      null;

    expect(() => assertFloorMatchesSeed(sheetConfigs, configs)).not.toThrow();
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

    expect(() => assertFloorMatchesSeed(sheetConfigs, configs)).not.toThrow();
  });
});
