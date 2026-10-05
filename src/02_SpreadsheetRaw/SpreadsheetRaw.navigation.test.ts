import { describe, expect, it } from "vitest";

import { type CellValue } from "../00_Source/CellValues/cellValues";
import { type RgbColor } from "../00_Source/RawSource/RgbColor";
import { stubSheetsService } from "../testSupport/fakeSheetsService";
import { assertType, type IsExactly } from "../testSupport/typeAssertions";
import { type CellRaw } from "./CellRaw";
import { type RowCommonRaw } from "./ClassBases/RowCommonRaw";
import { type CellFill, type CellStateRaw } from "./ClassTypes/StateRaw";
import { ColumnMetaRaw } from "./ColumnMetaRaw";
import { ColumnRaw } from "./ColumnRaw";
import { RowRaw } from "./RowRaw";
import { SheetMetaRaw } from "./SheetMetaRaw";
import { SpreadsheetRaw } from "./SpreadsheetRaw";
import { TableRaw } from "./TableRaw";
import { UniformRowRaw } from "./UniformRowRaw";

describe("SpreadsheetRaw navigation", () => {
  it("gives each accessor the class its return type names", () => {
    stubSheetsService({ sheets: [{ sheetId: 111, title: "Task Generic" }] });
    const raw = SpreadsheetRaw.init();
    const sheet = raw.sheet(111);
    const sheetMeta = raw.sheetMeta(111);
    const column = sheet.column(0);
    const columnMeta = sheetMeta.column(0);

    assertType<IsExactly<typeof sheet, TableRaw>>(true);
    assertType<IsExactly<typeof sheetMeta, SheetMetaRaw>>(true);
    assertType<IsExactly<typeof sheet.meta, SheetMetaRaw>>(true);
    assertType<IsExactly<typeof sheetMeta.primary, TableRaw>>(true);
    assertType<IsExactly<typeof column, ColumnRaw>>(true);
    assertType<IsExactly<typeof columnMeta, ColumnMetaRaw>>(true);
    assertType<IsExactly<typeof column.table, TableRaw>>(true);
    assertType<IsExactly<typeof columnMeta.table, SheetMetaRaw>>(true);
    assertType<IsExactly<typeof column.meta, ColumnMetaRaw>>(true);
    assertType<IsExactly<typeof columnMeta.primary, ColumnRaw>>(true);
    assertType<IsExactly<ReturnType<typeof sheet.row>, RowRaw>>(true);
    assertType<IsExactly<ReturnType<typeof sheet.rowCommon>, RowCommonRaw>>(
      true,
    );

    expect(sheet.meta).toBeInstanceOf(SheetMetaRaw);
    expect(sheetMeta.primary).toBeInstanceOf(TableRaw);
    expect(column).toBeInstanceOf(ColumnRaw);
    expect(columnMeta).toBeInstanceOf(ColumnMetaRaw);
    expect(column.table).toBeInstanceOf(TableRaw);
    expect(columnMeta.table).toBeInstanceOf(SheetMetaRaw);
    expect(column.meta).toBeInstanceOf(ColumnMetaRaw);
    expect(columnMeta.primary).toBeInstanceOf(ColumnRaw);
    expect(sheet.row(0)).toBeInstanceOf(RowRaw);
    expect(sheet.rowCommon(0)).toBeInstanceOf(RowRaw);
    expect(sheet.rowCommon(-4)).toBeInstanceOf(UniformRowRaw);
  });
});

describe("Raw value types", () => {
  it("declares CellStateRaw as the cell value and nothing else", () => {
    assertType<IsExactly<CellStateRaw, { value: CellValue }>>(true);
  });

  it("declares the blank the wire can hold, with nothing validating it away", () => {
    assertType<
      IsExactly<ReturnType<CellRaw<"boolean">["valueOrEmpty"]>, boolean | "">
    >(true);
    assertType<IsExactly<ReturnType<RowRaw["valueOrEmpty"]>, CellValue | "">>(
      true,
    );
    assertType<
      IsExactly<
        ReturnType<UniformRowRaw<"action">["valueOrEmpty"]>,
        boolean | ""
      >
    >(true);
    assertType<IsExactly<CellFill["backgroundColor"], RgbColor | undefined>>(
      true,
    );
  });
});
