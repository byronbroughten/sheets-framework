import { describe, expect, it } from "vitest";

import {
  configSheetFloorSeed,
  type FloorSeedColumn,
} from "./configSheetFloorSeed";

function seedColumn(
  sheetName: "spreadsheetConfig" | "sheetConfig" | "columnConfig",
  header: string,
): FloorSeedColumn {
  const column = configSheetFloorSeed[sheetName].columns.find(
    (entry) => entry.header === header,
  );
  if (column === undefined) {
    throw new Error(`Floor seed has no ${sheetName} column ${header}.`);
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

describe("configSheetFloorSeed column types", () => {
  it("declares TEXT, DOUBLE and BOOLEAN on the floor columns the spec names", () => {
    expect(seedColumn("spreadsheetConfig", "Table menu space").columnType).toBe(
      "TEXT",
    );
    expect(seedColumn("spreadsheetConfig", "ID header").columnType).toBe(
      "TEXT",
    );
    expect(seedColumn("spreadsheetConfig", "ID delimiter").columnType).toBe(
      "TEXT",
    );
    expect(
      seedColumn("spreadsheetConfig", "Start table column index base 1")
        .columnType,
    ).toBe("DOUBLE");
    expect(
      seedColumn("spreadsheetConfig", "Column ID row index base 1").columnType,
    ).toBe("DOUBLE");
    expect(
      seedColumn("spreadsheetConfig", "Column group heading row index base 1")
        .columnType,
    ).toBe("DOUBLE");
    expect(
      seedColumn("spreadsheetConfig", "Action row index base 1").columnType,
    ).toBe("DOUBLE");
    expect(
      seedColumn("spreadsheetConfig", "Table header row index base 1")
        .columnType,
    ).toBe("DOUBLE");
    expect(seedFeedback("Fill row IDs, time last ran").columnType).toBe("TEXT");
    expect(seedFeedback("Fill row IDs, run status").columnType).toBe("TEXT");
    expect(
      seedFeedback("Sync config sheet rows, time last ran").columnType,
    ).toBe("TEXT");
    expect(seedFeedback("Sync config sheet rows, run status").columnType).toBe(
      "TEXT",
    );

    expect(seedColumn("sheetConfig", "Sheet GID").columnType).toBe("DOUBLE");
    expect(seedColumn("sheetConfig", "Sheet title").columnType).toBe("TEXT");
    expect(seedColumn("sheetConfig", "Let api access").columnType).toBe(
      "BOOLEAN",
    );

    expect(seedColumn("columnConfig", "Sheet GID").columnType).toBe("DOUBLE");
    expect(seedColumn("columnConfig", "Column ID").columnType).toBe("TEXT");
    expect(seedColumn("columnConfig", "Sheet title").columnType).toBe("TEXT");
    expect(seedColumn("columnConfig", "Header").columnType).toBe("TEXT");
    expect(seedColumn("columnConfig", "Empty value allowed").columnType).toBe(
      "BOOLEAN",
    );
  });
});

describe("configSheetFloorSeed declared cells", () => {
  it("declares Let api access true per floor tab and Empty value allowed false per floor column, endpoints included", () => {
    expect(configSheetFloorSeed.spreadsheetConfig.letApiAccess).toBe(true);
    expect(configSheetFloorSeed.sheetConfig.letApiAccess).toBe(true);
    expect(configSheetFloorSeed.columnConfig.letApiAccess).toBe(true);
    expect(configSheetFloorSeed.valueConfig.letApiAccess).toBe(true);

    const floorColumns = [
      ...configSheetFloorSeed.spreadsheetConfig.columns,
      ...Object.values(
        configSheetFloorSeed.spreadsheetConfig.endpoints,
      ).flatMap((endpoint) => [endpoint.timeLastRan, endpoint.runStatus]),
      ...configSheetFloorSeed.sheetConfig.columns,
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
      ...configSheetFloorSeed.sheetConfig.columns,
      ...configSheetFloorSeed.columnConfig.columns,
    ];
    otherColumns.forEach((column) => {
      expect(column).not.toHaveProperty("dataValue");
    });
  });
});
