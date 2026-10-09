import { describe, expect, it } from "vitest";

import {
  configSheetFloorSeed,
  type FloorSeedColumn,
  floorSeedColumns,
} from "./configSheetFloorSeed";

function seedColumn(
  tableName: "spreadsheetConfig" | "tableConfig" | "columnConfig",
  header: string,
): FloorSeedColumn {
  const column = configSheetFloorSeed[tableName].columns.find(
    (entry) => entry.header === header,
  );
  if (column === undefined) {
    throw new Error(`Floor seed has no ${tableName} column ${header}.`);
  }
  return column;
}

function seedFeedback(header: string) {
  const column = Object.values(configSheetFloorSeed.spreadsheetConfig.endpoints)
    .flatMap((endpoint) => [endpoint.timeLastRan, endpoint.runStatus])
    .find((entry) => entry.header === header);
  if (column === undefined) {
    throw new Error(`Floor seed has no Spreadsheet Config feedback ${header}.`);
  }
  return column;
}

describe("configSheetFloorSeed Spreadsheet Config", () => {
  it("declares only Table menu space and the framework endpoint columns", () => {
    expect(
      floorSeedColumns("spreadsheetConfig").map((column) => column.header),
    ).toEqual([
      "Table menu space",
      "Fill row IDs, time last ran",
      "Fill row IDs, run status",
      "Sync config sheet rows, time last ran",
      "Sync config sheet rows, run status",
    ]);
  });
});

describe("configSheetFloorSeed column types", () => {
  it("declares TEXT and BOOLEAN on the floor columns the spec names", () => {
    expect(seedColumn("spreadsheetConfig", "Table menu space").columnType).toBe(
      "TEXT",
    );
    expect(seedFeedback("Fill row IDs, time last ran").columnType).toBe("TEXT");
    expect(seedFeedback("Fill row IDs, run status").columnType).toBe("TEXT");
    expect(
      seedFeedback("Sync config sheet rows, time last ran").columnType,
    ).toBe("TEXT");
    expect(seedFeedback("Sync config sheet rows, run status").columnType).toBe(
      "TEXT",
    );

    expect(seedColumn("tableConfig", "Table ID").columnType).toBe("TEXT");
    expect(seedColumn("tableConfig", "Table name").columnType).toBe("TEXT");
    expect(seedColumn("tableConfig", "Sheet title").columnType).toBe("TEXT");
    expect(seedColumn("tableConfig", "Let api access").columnType).toBe(
      "BOOLEAN",
    );

    expect(seedColumn("columnConfig", "Table ID").columnType).toBe("TEXT");
    expect(seedColumn("columnConfig", "Column ID").columnType).toBe("TEXT");
    expect(seedColumn("columnConfig", "Table name").columnType).toBe("TEXT");
    expect(seedColumn("columnConfig", "Header").columnType).toBe("TEXT");
    expect(seedColumn("columnConfig", "Empty value allowed").columnType).toBe(
      "BOOLEAN",
    );
  });
});

describe("configSheetFloorSeed declared cells", () => {
  it("declares Let api access true per floor tab and Empty value allowed false per floor column, endpoints included", () => {
    expect(configSheetFloorSeed.spreadsheetConfig.letApiAccess).toBe(true);
    expect(configSheetFloorSeed.tableConfig.letApiAccess).toBe(true);
    expect(configSheetFloorSeed.columnConfig.letApiAccess).toBe(true);
    expect(configSheetFloorSeed.valueConfig.letApiAccess).toBe(true);

    const floorColumns = [
      ...configSheetFloorSeed.spreadsheetConfig.columns,
      ...Object.values(
        configSheetFloorSeed.spreadsheetConfig.endpoints,
      ).flatMap((endpoint) => [endpoint.timeLastRan, endpoint.runStatus]),
      ...configSheetFloorSeed.tableConfig.columns,
      ...configSheetFloorSeed.columnConfig.columns,
    ];
    expect(floorColumns.length).toBeGreaterThan(0);
    floorColumns.forEach((column) => {
      expect(column.emptyValueAllowed).toBe(false);
    });
  });
});

describe("configSheetFloorSeed data values", () => {
  it("declares Not used as Table menu space's data value", () => {
    expect(seedColumn("spreadsheetConfig", "Table menu space").dataValue).toBe(
      "Not used",
    );
  });

  it("declares no data value on any other floor column", () => {
    const otherColumns = [
      ...configSheetFloorSeed.spreadsheetConfig.columns.filter(
        (column) => column.header !== "Table menu space",
      ),
      ...configSheetFloorSeed.tableConfig.columns,
      ...configSheetFloorSeed.columnConfig.columns,
    ];
    otherColumns.forEach((column) => {
      expect(column).not.toHaveProperty("dataValue");
    });
  });
});
