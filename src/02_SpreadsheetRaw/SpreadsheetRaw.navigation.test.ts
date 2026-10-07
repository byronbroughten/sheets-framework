import { describe, expect, it } from "vitest";

import { type CellValue } from "../00_Source/CellValues/cellValues";
import { type RgbColor } from "../00_Source/RawSource/RgbColor";
import { type HeadRowValueName } from "../01_SpreadsheetSchema/headRows";
import { stubSheetsService } from "../testSupport/fakeSheetsService";
import { assertType, type IsExactly } from "../testSupport/typeAssertions";
import { CellRaw } from "./CellRaw";
import { type ColumnBaseRaw } from "./ClassBases/ColumnBaseRaw";
import { type RowCommonRaw } from "./ClassBases/RowCommonRaw";
import { type TableBaseRaw } from "./ClassBases/TableBaseRaw";
import { type CellFill, type CellStateRaw } from "./ClassTypes/StateRaw";
import { ColumnProfileRaw } from "./ColumnProfileRaw";
import { ColumnRaw } from "./ColumnRaw";
import { HeadRowRaw } from "./HeadRowRaw";
import { RowRaw } from "./RowRaw";
import { SpreadsheetRaw } from "./SpreadsheetRaw";
import { placedTableSheet, tableId111 } from "./spreadsheetRawTestSupport";
import { TableProfileRaw } from "./TableProfileRaw";
import { TableRaw } from "./TableRaw";

describe("SpreadsheetRaw navigation", () => {
  it("gives each accessor the class its return type names", () => {
    stubSheetsService({
      sheets: [placedTableSheet({ sheetId: 111, title: "Task Generic" })],
    });
    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    const table = raw.table(tableId111);
    const tableOnSheet = raw.tableOnSheet(111);
    const column = table.column(0);
    const headRow = table.headRow("action");
    const headCell = column.headCell("header");

    assertType<IsExactly<typeof table, TableRaw>>(true);
    assertType<IsExactly<typeof tableOnSheet, TableRaw>>(true);
    assertType<IsExactly<typeof column, ColumnRaw>>(true);
    assertType<IsExactly<typeof column.table, TableRaw>>(true);
    assertType<IsExactly<typeof table.profile, TableProfileRaw>>(true);
    assertType<IsExactly<typeof column.profile, ColumnProfileRaw>>(true);
    assertType<
      IsExactly<ReturnType<typeof table.profile.columnById>, ColumnProfileRaw>
    >(true);
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
    expect(tableOnSheet).toBeInstanceOf(TableRaw);
    expect(column).toBeInstanceOf(ColumnRaw);
    expect(column.table).toBeInstanceOf(TableRaw);
    expect(table.profile).toBeInstanceOf(TableProfileRaw);
    expect(column.profile).toBeInstanceOf(ColumnProfileRaw);
    expect(table.row(0)).toBeInstanceOf(RowRaw);
    expect(table.rowCommon(0)).toBeInstanceOf(RowRaw);
    expect(table.rowCommon(-4)).toBeInstanceOf(HeadRowRaw);
    expect(headRow).toBeInstanceOf(HeadRowRaw);
    expect(headRow.table).toBeInstanceOf(TableRaw);
    expect(table.headRowByIndex(-2)).toBeInstanceOf(HeadRowRaw);
    expect(headCell).toBeInstanceOf(CellRaw);
  });

  it("offers no Meta view", () => {
    assertType<
      IsExactly<
        Extract<
          "meta" | "sheetMeta",
          keyof SpreadsheetRaw | keyof TableRaw | keyof ColumnRaw
        >,
        never
      >
    >(true);
  });
});

describe("Raw profile members", () => {
  it("gives the Table profile its column IDs, ID prefix and lookup by ID", () => {
    assertType<
      IsExactly<
        Exclude<keyof TableProfileRaw, keyof TableBaseRaw>,
        "columnIds" | "idPrefix" | "columnById"
      >
    >(true);
  });

  it("gives the column profile its descriptive facts and value titles", () => {
    assertType<
      IsExactly<
        Exclude<keyof ColumnProfileRaw, keyof ColumnBaseRaw>,
        | "columnType"
        | "isFormula"
        | "numberFormatType"
        | "topValue"
        | "dataValidationConditionType"
        | "header"
        | "valueValidationStrings"
        | "validationConditionType"
        | "valueTitle"
        | "declaredValueTitle"
        | "validationValueTitle"
      >
    >(true);
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
      IsExactly<ReturnType<HeadRowRaw<"columnId">["valueOrEmpty"]>, string | "">
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
