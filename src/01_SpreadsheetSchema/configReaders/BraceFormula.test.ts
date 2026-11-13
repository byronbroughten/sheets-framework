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
  it("wraps an @ brace in SINGLE, as this row's value, of its own Table or another", () => {
    expect(
      BraceFormula.compileForCell({
        configs,
        tableKey: "item",
        formula: "=${@price}*${@count}+${unit[@rent]}",
      }),
    ).toBe(
      "=SINGLE(Items[Purchase Price])*SINGLE(Items[Count])+SINGLE(Unit_list[Rent])",
    );
  });

  it("gives a brace with no @ the bare reference, the whole column", () => {
    expect(
      BraceFormula.compileForCell({
        configs,
        tableKey: "item",
        formula: "=COUNTA(${unit[rent]})+SUM(${item[count]})+SUM(${price})",
      }),
    ).toBe(
      "=COUNTA(Unit_list[Rent])+SUM(Items[Count])+SUM(Items[Purchase Price])",
    );
  });

  it("leaves a brace inside a string literal as text", () => {
    expect(
      BraceFormula.compileForCell({
        configs,
        tableKey: "item",
        formula: '="${@count} is ""${count}"""&${@count}',
      }),
    ).toBe('="${@count} is ""${count}"""&SINGLE(Items[Count])');
  });

  it("refuses a brace naming no column, naming the brace and the formula", () => {
    expect(() =>
      BraceFormula.compileForCell({
        configs,
        tableKey: "item",
        formula: "=${@prise}*2",
      }),
    ).toThrow(
      "The brace ${@prise} names no column of Table item, in formula =${@prise}*2.",
    );
    expect(() =>
      BraceFormula.compileForCell({
        configs,
        tableKey: "item",
        formula: "=${unit[@price]}",
      }),
    ).toThrow(
      "The brace ${unit[@price]} names no column of Table unit, in formula =${unit[@price]}.",
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
  it("turns every SINGLE into an @ brace, bare for its own Table's columns and qualified for another's", () => {
    expect(
      BraceFormula.decompile({
        configs,
        tableKey: "item",
        formula:
          "=SINGLE(Items[Purchase Price])*COUNTA(Unit_list[Rent])+SUM(Items[Count])+SINGLE(Unit_list[Rent])",
      }),
    ).toBe("=${@price}*COUNTA(${unit[rent]})+SUM(${count})+${unit[@rent]}");
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
    ).toBe('=Other[Thing]&Items[Gone]&"Items[Count]"&${@count}');
  });

  it("gives back the brace form that compiled to the live text", () => {
    const formula =
      "=IF(${@count}>0, ${@price}/COUNTA(${unit[rent]}), ${unit[@rent]}+SUM(${count}))";
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

  it("anchors an @ brace's column to the first body row, relative by row", () => {
    expect(
      BraceFormula.compileForRule({
        tableKey: "item",
        layout,
        formula: "=${@price}>${item[@newColumn]}",
      }),
    ).toBe("=$E5>$D5");
  });

  it("gives a brace with no @ the whole body column", () => {
    expect(
      BraceFormula.compileForRule({
        tableKey: "item",
        layout,
        formula: "=COUNTIF(${count}, ${@count})>COUNTA(${item[price]})",
      }),
    ).toBe("=COUNTIF($C$5:$C, $C5)>COUNTA($E$5:$E)");
  });

  it("refuses a brace naming no column of the layout, or another Table", () => {
    expect(() =>
      BraceFormula.compileForRule({
        tableKey: "item",
        layout,
        formula: "=${@prise}",
      }),
    ).toThrow(
      "The brace ${@prise} names no column of Table item, in formula =${@prise}.",
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

describe("BraceFormula.decompileForRule", () => {
  // Header row 4 (index 3) from column C, so the first body row is row 5.
  const layout = {
    headerTopLeft: { rowIndex: SheetIndex.row(3), colIndex: SheetIndex.col(2) },
    columnNamesLeftToRight: ["count", "newColumn", "price"],
  };

  function decompile(formula: string, isOneColumn = false) {
    return BraceFormula.decompileForRule({
      layout,
      formula,
      isOneColumn,
    });
  }

  it("gives back the brace form a rule's formula compiled from", () => {
    expect(decompile("=COUNTIF($C$5:$C, $C5)>1")).toBe(
      "=COUNTIF(${count}, ${@count})>1",
    );
    expect(decompile('=$E5>$D5&"$C5"')).toBe('=${@price}>${@newColumn}&"$C5"');
  });

  it("reads a column-relative reference as its column only on a one-column rule", () => {
    expect(decompile("=D5>2", true)).toBe("=${@newColumn}>2");
    expect(decompile("=D5>2")).toBeUndefined();
  });

  it("can't express a reference off the first body row, outside the Table or on another tab", () => {
    expect(decompile("=$D6>2")).toBeUndefined();
    expect(decompile("=$D$5>2")).toBeUndefined();
    expect(decompile("=$B5>2")).toBeUndefined();
    expect(decompile("=Other!$D5>2")).toBeUndefined();
    expect(decompile("=COUNTA($D:$D)>2")).toBeUndefined();
  });

  it("leaves function names alone", () => {
    expect(decompile("=LOG10($E5)>ATAN2($C5, 1)")).toBe(
      "=LOG10(${@price})>ATAN2(${@count}, 1)",
    );
  });
});

describe("BraceFormula.malformedBraces", () => {
  it("names each ${…} that reads as no brace, outside string literals", () => {
    expect(
      BraceFormula.malformedBraces({
        formula: '=${@vendor[cost]}+${@price}+${unit[@rent]}+${@@x}&"${@a[b]}"',
      }),
    ).toEqual(["${@vendor[cost]}", "${@@x}"]);
  });
});

describe("BraceFormula.columnsNamed", () => {
  it("names the column of every brace form, its own Table's and another's", () => {
    expect(
      BraceFormula.columnsNamed({
        tableKey: "item",
        formula:
          '=${@price}+${count}+${unit[@rent]}+SUM(${unit[rent]})&"${gone}"',
      }),
    ).toEqual([
      { brace: "${@price}", tableKey: "item", columnName: "price" },
      { brace: "${count}", tableKey: "item", columnName: "count" },
      { brace: "${unit[@rent]}", tableKey: "unit", columnName: "rent" },
      { brace: "${unit[rent]}", tableKey: "unit", columnName: "rent" },
    ]);
  });
});
