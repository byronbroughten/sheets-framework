import { describe, expect, it } from "vitest";

import { SheetIndex } from "../../00_Source/RawSource/SheetIndex";
import { type BraceConfigs, BraceFormula } from "./BraceFormula";

function column(columnId: string, header: string) {
  return {
    columnId,
    header,
    valueName: "number",
    isFormula: false,
    emptyValueAllowed: true,
    customDefaultValue: null,
  };
}

const itemTable = {
  tableId: "tbl-item",
  tableName: "Items",
  sheetGid: 1,
  idPrefix: "itm",
  hasIdColumn: true,
  hasNameColumn: false,
};
const itemColumns = {
  price: column("c:itm:price", "Purchase Price"),
  count: column("c:itm:count", "Count"),
};

const configs: BraceConfigs = {
  tableConfigs: {
    item: itemTable,
    unit: {
      tableId: "tbl-unit",
      tableName: "Unit_list",
      sheetGid: 2,
      idPrefix: "unt",
      hasIdColumn: true,
      hasNameColumn: false,
    },
  },
  columnConfigs: {
    item: itemColumns,
    unit: {
      rent: column("c:unt:rent", "Rent"),
    },
  },
};

describe("BraceFormula.compileForCell", () => {
  it("wraps a column of the formula's own Table in SINGLE, as its same-row value", () => {
    expect(
      BraceFormula.compileForCell({
        configs,
        tableKey: "item",
        formula: "=${price}*${count}",
      }),
    ).toBe("=SINGLE(Items[Purchase Price])*SINGLE(Items[Count])");
  });

  it("gives a Table-qualified brace the bare reference, the whole column", () => {
    expect(
      BraceFormula.compileForCell({
        configs,
        tableKey: "item",
        formula: "=COUNTA(${unit[rent]})+SUM(${item[count]})",
      }),
    ).toBe("=COUNTA(Unit_list[Rent])+SUM(Items[Count])");
  });

  it("leaves a brace inside a string literal as text", () => {
    expect(
      BraceFormula.compileForCell({
        configs,
        tableKey: "item",
        formula: '="${count} is ""${count}"""&${count}',
      }),
    ).toBe('="${count} is ""${count}"""&SINGLE(Items[Count])');
  });

  it("refuses a brace naming no column, naming the brace and the formula", () => {
    expect(() =>
      BraceFormula.compileForCell({
        configs,
        tableKey: "item",
        formula: "=${prise}*2",
      }),
    ).toThrow(
      "The brace ${prise} names no column of Table item, in formula =${prise}*2.",
    );
    expect(() =>
      BraceFormula.compileForCell({
        configs,
        tableKey: "item",
        formula: "=${unit[price]}",
      }),
    ).toThrow(
      "The brace ${unit[price]} names no column of Table unit, in formula =${unit[price]}.",
    );
    expect(() =>
      BraceFormula.compileForCell({
        configs,
        tableKey: "item",
        formula: "=${units[rent]}",
      }),
    ).toThrow(
      "The brace ${units[rent]} names no Table, in formula =${units[rent]}.",
    );
  });
});

describe("BraceFormula.decompile", () => {
  it("turns an own-Table SINGLE into a bare brace and any other reference into a qualified one", () => {
    expect(
      BraceFormula.decompile({
        configs,
        tableKey: "item",
        formula:
          "=SINGLE(Items[Purchase Price])*COUNTA(Unit_list[Rent])+SUM(Items[Count])",
      }),
    ).toBe("=${price}*COUNTA(${unit[rent]})+SUM(${item[count]})");
  });

  it("keeps SINGLE as text around another Table's column", () => {
    expect(
      BraceFormula.decompile({
        configs,
        tableKey: "item",
        formula: "=SINGLE(Unit_list[Rent])",
      }),
    ).toBe("=SINGLE(${unit[rent]})");
  });

  it("reads a Table and header rename that Sheets propagated as the same brace form", () => {
    const renamed: BraceConfigs = {
      ...configs,
      tableConfigs: {
        ...configs.tableConfigs,
        item: { ...itemTable, tableName: "Inventory" },
      },
      columnConfigs: {
        ...configs.columnConfigs,
        item: {
          ...itemColumns,
          price: column("c:itm:price", "Price paid"),
        },
      },
    };
    const before = "=SINGLE(Items[Purchase Price])*2";
    const after = "=SINGLE(Inventory[Price paid])*2";
    expect(
      BraceFormula.decompile({
        configs: renamed,
        tableKey: "item",
        formula: after,
      }),
    ).toBe(
      BraceFormula.decompile({ configs, tableKey: "item", formula: before }),
    );
  });

  it("leaves a reference to no configured column, and string literals, as they are", () => {
    expect(
      BraceFormula.decompile({
        configs,
        tableKey: "item",
        formula:
          '=Other[Thing]&Items[Gone]&"Items[Count]"&SINGLE(\nItems[Count])',
      }),
    ).toBe('=Other[Thing]&Items[Gone]&"Items[Count]"&${count}');
  });

  it("gives back the brace form that compiled to the live text", () => {
    const formula =
      "=IF(${count}>0, ${price}/COUNTA(${unit[rent]}), SINGLE(${unit[rent]}))";
    const compiled = BraceFormula.compileForCell({
      configs,
      tableKey: "item",
      formula,
    });
    expect(
      BraceFormula.decompile({ configs, tableKey: "item", formula: compiled }),
    ).toBe(formula);
  });
});

describe("BraceFormula.compileForRule", () => {
  // Header row 4 (index 3) from column C, so the first body row is row 5.
  const layout = {
    headerTopLeft: { rowIndex: SheetIndex.row(3), colIndex: SheetIndex.col(2) },
    columnNamesLeftToRight: ["count", "newColumn", "price"],
  };

  it("anchors a bare brace's column to the first body row, relative by row", () => {
    expect(
      BraceFormula.compileForRule({
        tableKey: "item",
        layout,
        formula: "=${price}>${newColumn}",
      }),
    ).toBe("=$E5>$D5");
  });

  it("gives an own-Table qualified brace the whole body column", () => {
    expect(
      BraceFormula.compileForRule({
        tableKey: "item",
        layout,
        formula: "=COUNTIF(${item[count]}, ${count})>1",
      }),
    ).toBe("=COUNTIF($C$5:$C, $C5)>1");
  });

  it("refuses a brace naming no column of the layout, or another Table", () => {
    expect(() =>
      BraceFormula.compileForRule({
        tableKey: "item",
        layout,
        formula: "=${prise}",
      }),
    ).toThrow(
      "The brace ${prise} names no column of Table item, in formula =${prise}.",
    );
    expect(() =>
      BraceFormula.compileForRule({
        tableKey: "item",
        layout,
        formula: "=${unit[rent]}",
      }),
    ).toThrow(
      "The brace ${unit[rent]} names another Table, which a rule's formula can't reach, in formula =${unit[rent]}.",
    );
  });
});
