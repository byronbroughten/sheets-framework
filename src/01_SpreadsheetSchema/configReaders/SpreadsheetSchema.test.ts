import { describe, expect, it } from "vitest";

import { SheetIndex } from "../../00_Source/RawSource/SheetIndex";
import { assertType, type IsExactly } from "../../testSupport/typeAssertions";
import {
  type ColumnFullName,
  type ColumnName,
  type ColumnNameFiltered,
  type ColumnNameOf,
  type ColumnValue,
  type ColumnValueName,
  getColumnTraitByName,
  type MakeColumnFullName,
  type SheetNameOf,
  type ValueNameOf,
  type ValueOf,
} from "./columnConfigsTypes";
import { ColumnSchema } from "./ColumnSchema";
import { SpreadsheetSchema } from "./SpreadsheetSchema";
import { tableKeysByGid, type TableName } from "./tableConfigsTypes";
import { TableSchema } from "./TableSchema";
import type { ValueName } from "./valueSchemas";

describe("SpreadsheetSchema", () => {
  const schema = new SpreadsheetSchema();

  describe("combineNames", () => {
    it("joins two names with the name delimiter", () => {
      expect(schema.combineNames("foo", "bar")).toBe("foo_bar");
    });
  });

  describe("table placement", () => {
    const runItem = schema.sheetByName("runItem");

    it("maps a sheet gid to its managed Tables, and an unmanaged one to none", () => {
      expect(
        schema.tablesOnGid(runItem.sheetGid).map((table) => table.tableName),
      ).toEqual(["runItem"]);
      expect(schema.tablesOnGid(-1)).toEqual([]);
    });

    it("labels a position in the numbering Sheets shows the operator", () => {
      const row = SheetIndex.row;
      const col = SheetIndex.col;
      expect(schema.positionLabel(row(0), col(0))).toBe("row 1, column A");
      expect(schema.positionLabel(row(3), col(1))).toBe("row 4, column B");
      expect(schema.positionLabel(row(3), col(26))).toBe("row 4, column AA");
    });
  });

  describe("isInSheetGids", () => {
    it("agrees with the generated sheet gid list", () => {
      const [firstGid] = tableKeysByGid().keys();
      expect(firstGid).toBeDefined();
      expect(schema.isInSheetGids(firstGid as number)).toBe(true);
      expect(schema.isInSheetGids(Number.MAX_SAFE_INTEGER)).toBe(false);
    });
  });

  describe("config", () => {
    it("exposes the hand-authored spreadsheet layout constants", () => {
      expect(schema.idDelimiter).toBe(":");
      expect(schema.colIdRowIndex).toBe(-4);
      expect(schema.actionRowIndex).toBe(-2);
      expect(schema.tableHeaderRowIndex).toBe(-1);
    });
  });
});

describe("type-level precision", () => {
  it("resolves a name-addressed column to its exact literal types", () => {
    const column = ColumnSchema.fromColumnName("item", "requiredCount");
    assertType<IsExactly<typeof column.valueName, "number">>(true);
    assertType<IsExactly<typeof column.columnName, "requiredCount">>(true);
    assertType<IsExactly<typeof column.tableName, "item">>(true);
    assertType<IsExactly<typeof column.fullName, "item_requiredCount">>(true);
    assertType<
      IsExactly<ReturnType<typeof column.makeDefaultDataValue>, number | "">
    >(true);
    assertType<IsExactly<ColumnValue<"item", "requiredCount">, number | "">>(
      true,
    );
    expect(column.valueName).toBe("number");
    expect(column.fullName).toBe("item_requiredCount");
  });

  it("resolves a columnId-addressed column to the usable widened types, never `never`", () => {
    const column = ColumnSchema.fromColumnId(
      "item",
      getColumnTraitByName("item", "requiredCount", "columnId"),
    );
    assertType<IsExactly<typeof column.valueName, ValueName>>(true);
    assertType<IsExactly<typeof column.columnName, ColumnName<TableName>>>(
      true,
    );
    assertType<IsExactly<typeof column.fullName, ColumnFullName>>(true);
    expect(column.valueName).toBe("number");
    expect(column.columnName).toBe("requiredCount");
  });

  it("keeps the sheet trait accessor's shape at both instantiations", () => {
    const byName = TableSchema.fromSheetName("item");
    const byGid = TableSchema.fromSheetGid(byName.sheetGid);
    assertType<IsExactly<typeof byName.tableName, "item">>(true);
    assertType<IsExactly<typeof byGid.tableName, TableName>>(true);
    assertType<
      IsExactly<ReturnType<typeof byName.trait<"hasIdColumn">>, boolean>
    >(true);
    assertType<IsExactly<typeof byName.columnNames, ColumnName<"item">[]>>(
      true,
    );
    expect(byGid.tableName).toBe("item");
    expect(byName.columnNames).toContain("requiredCount");
  });

  it("reaches a sheet from the spreadsheet schema by either address", () => {
    const ss = new SpreadsheetSchema();
    const byName = ss.sheetByName("item");
    const byGid = ss.sheetByGid(byName.sheetGid);
    assertType<IsExactly<typeof byName, TableSchema<"item">>>(true);
    assertType<IsExactly<typeof byGid, TableSchema>>(true);
    expect(byGid.sheetGid).toBe(byName.sheetGid);
  });

  it("navigates from a column schema back to its own sheet", () => {
    const sheet = ColumnSchema.fromColumnName("item", "requiredCount").sheet;
    assertType<IsExactly<typeof sheet, TableSchema<"item">>>(true);
    expect(sheet.tableName).toBe("item");
  });
});

function columnConfigOf(fullName: ColumnFullName) {
  const column = ColumnSchema.fromColumnName(
    ...(fullName.split("_") as [TableName, never]),
  );
  return column;
}
function valueNameOfFullName(fullName: ColumnFullName): ValueName {
  return columnConfigOf(fullName).valueName;
}
function columnFullNameIsFormula(fullName: ColumnFullName): boolean {
  return columnConfigOf(fullName).isFormula;
}

describe("ColumnFullName, absolute column addressing", () => {
  it("narrows to a proper subset of columns when filtered by value name", () => {
    const sampled: ColumnFullName<"boolean"> = "valueTypes_sampledBoolean";
    // @ts-expect-error a number column is not a boolean column
    const numeric: ColumnFullName<"boolean"> = "item_requiredCount";
    // @ts-expect-error a string column is not a boolean column
    const text: ColumnFullName<"boolean"> =
      "spreadsheetConfig_fillRowIdsTimeLastRan";
    expect(valueNameOfFullName(sampled)).toBe("boolean");
    expect(valueNameOfFullName(numeric)).toBe("number");
    expect(valueNameOfFullName(text)).toBe("string");
  });

  // The two names are siblings, not one inside the other.
  it("keeps a declared checkbox column out of the boolean set, and the reverse", () => {
    const declared: ColumnFullName<"checkbox"> = "valueTypes_checkbox";
    // @ts-expect-error a declared checkbox column is no longer a boolean column
    const asBoolean: ColumnFullName<"boolean"> = "valueTypes_checkbox";
    // @ts-expect-error a column that only samples as boolean is not a checkbox column
    const sampled: ColumnFullName<"checkbox"> = "valueTypes_sampledBoolean";
    expect(valueNameOfFullName(declared)).toBe("checkbox");
    expect(valueNameOfFullName(asBoolean)).toBe("checkbox");
    expect(valueNameOfFullName(sampled)).toBe("boolean");
  });

  it("denotes every column, formula ones included, when unfiltered", () => {
    // Derived through the relative family, so this is an independent check
    // that collapsing the old "simple" union lost nothing.
    type EveryColumnFullName = {
      [TN in TableName]: MakeColumnFullName<TN, ColumnName<TN>>;
    }[TableName];
    assertType<IsExactly<ColumnFullName, EveryColumnFullName>>(true);
  });

  it("narrows further on the formula axis", () => {
    const writable: ColumnFullName<"boolean", false> =
      "valueTypes_sampledBoolean";
    expect(columnFullNameIsFormula(writable)).toBe(false);
  });

  it("resolves a full name's sheet, column, value name and value type exactly", () => {
    type FN = "valueTypes_checkbox";
    assertType<IsExactly<SheetNameOf<FN>, "valueTypes">>(true);
    assertType<IsExactly<ColumnNameOf<FN>, "checkbox">>(true);
    assertType<IsExactly<ValueNameOf<FN>, "checkbox">>(true);
    assertType<IsExactly<ValueOf<FN>, boolean>>(true);
  });

  it("agrees with relative addressing on sampled columns", () => {
    assertType<
      IsExactly<
        ValueNameOf<"valueTypes_checkbox">,
        ColumnValueName<"valueTypes", "checkbox">
      >
    >(true);
    assertType<
      IsExactly<
        ValueOf<"item_requiredCount">,
        ColumnValue<"item", "requiredCount">
      >
    >(true);
    assertType<
      IsExactly<ColumnNameOf<"valueConfig_exampleValue">, "exampleValue">
    >(true);
  });

  it("filters a column name within a sheet on the same two axes", () => {
    assertType<
      IsExactly<
        ColumnNameFiltered<"tableConfig", "checkbox", false>,
        "letApiAccess"
      >
    >(true);
    assertType<IsExactly<ColumnNameFiltered<"item">, ColumnName<"item">>>(true);
  });
});
