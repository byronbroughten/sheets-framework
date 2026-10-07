import { describe, expect, it } from "vitest";

import { type CellValue } from "../00_Source/CellValues/cellValues";
import { type RgbColor } from "../00_Source/RawSource/RgbColor";
import { type HeadRowValueName } from "../01_SpreadsheetSchema/headRows";
import { stubSheetsService } from "../testSupport/fakeSheetsService";
import { assertType, type IsExactly } from "../testSupport/typeAssertions";
import { CellRaw } from "./CellRaw";
import { type RowCommonRaw } from "./ClassBases/RowCommonRaw";
import { type CellFill, type CellStateRaw } from "./ClassTypes/StateRaw";
import { ColumnMetaRaw } from "./ColumnMetaRaw";
import { ColumnRaw } from "./ColumnRaw";
import { HeadRowRaw } from "./HeadRowRaw";
import { RowRaw } from "./RowRaw";
import { SheetMetaRaw } from "./SheetMetaRaw";
import { SpreadsheetRaw } from "./SpreadsheetRaw";
import { placedTableSheet, tableId111 } from "./spreadsheetRawTestSupport";
import { TableRaw } from "./TableRaw";

describe("SpreadsheetRaw navigation", () => {
  it("gives each accessor the class its return type names", () => {
    stubSheetsService({
      sheets: [placedTableSheet({ sheetId: 111, title: "Task Generic" })],
    });
    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    const table = raw.table(tableId111);
    const sheetMeta = raw.sheetMeta(111);
    const column = table.column(0);
    const columnMeta = sheetMeta.column(0);
    const headRow = table.headRow("action");
    const headCell = column.headCell("header");

    assertType<IsExactly<typeof table, TableRaw>>(true);
    assertType<IsExactly<typeof sheetMeta, SheetMetaRaw>>(true);
    assertType<IsExactly<typeof table.meta, SheetMetaRaw>>(true);
    assertType<IsExactly<typeof sheetMeta.primary, TableRaw>>(true);
    assertType<IsExactly<typeof column, ColumnRaw>>(true);
    assertType<IsExactly<typeof columnMeta, ColumnMetaRaw>>(true);
    assertType<IsExactly<typeof column.table, TableRaw>>(true);
    assertType<IsExactly<typeof columnMeta.table, SheetMetaRaw>>(true);
    assertType<IsExactly<typeof column.meta, ColumnMetaRaw>>(true);
    assertType<IsExactly<typeof columnMeta.primary, ColumnRaw>>(true);
    assertType<IsExactly<ReturnType<typeof table.row>, RowRaw>>(true);
    assertType<IsExactly<ReturnType<typeof table.rowCommon>, RowCommonRaw>>(
      true,
    );
    assertType<IsExactly<typeof headRow, HeadRowRaw<"action">>>(true);
    assertType<IsExactly<typeof headRow.table, TableRaw>>(true);
    assertType<IsExactly<ReturnType<typeof table.headRowByIndex>, HeadRowRaw>>(
      true,
    );
    assertType<IsExactly<typeof headCell, CellRaw<"string">>>(true);

    expect(table).toBeInstanceOf(TableRaw);
    expect(table.meta).toBeInstanceOf(SheetMetaRaw);
    expect(sheetMeta.primary).toBeInstanceOf(TableRaw);
    expect(column).toBeInstanceOf(ColumnRaw);
    expect(columnMeta).toBeInstanceOf(ColumnMetaRaw);
    expect(column.table).toBeInstanceOf(TableRaw);
    expect(columnMeta.table).toBeInstanceOf(SheetMetaRaw);
    expect(column.meta).toBeInstanceOf(ColumnMetaRaw);
    expect(columnMeta.primary).toBeInstanceOf(ColumnRaw);
    expect(table.row(0)).toBeInstanceOf(RowRaw);
    expect(table.rowCommon(0)).toBeInstanceOf(RowRaw);
    expect(table.rowCommon(-4)).toBeInstanceOf(HeadRowRaw);
    expect(headRow).toBeInstanceOf(HeadRowRaw);
    expect(headRow.table).toBeInstanceOf(TableRaw);
    expect(table.headRowByIndex(-2)).toBeInstanceOf(HeadRowRaw);
    expect(headCell).toBeInstanceOf(CellRaw);
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
        ReturnType<HeadRowRaw<"columnId">["valueOrEmpty"]>,
        string | ""
      >
    >(true);
    assertType<IsExactly<CellFill["backgroundColor"], RgbColor | undefined>>(
      true,
    );
  });

  it("types a head row's cells as the union of its roles' values", () => {
    assertType<
      IsExactly<
        ReturnType<HeadRowRaw<"action">["valueOrEmpty"]>,
        boolean | string
      >
    >(true);
    assertType<
      IsExactly<
        ReturnType<HeadRowRaw<"groupHeading2">["valueOrEmpty"]>,
        boolean | string
      >
    >(true);
    assertType<
      IsExactly<ReturnType<HeadRowRaw<"header">["valueOrEmpty"]>, string>
    >(true);
    assertType<
      IsExactly<HeadRowRaw<"action">["roles"], ("action" | "groupHeading2")[]>
    >(true);
    assertType<
      IsExactly<HeadRowValueName<"groupHeading2">, "boolean" | "string">
    >(true);
  });
});
