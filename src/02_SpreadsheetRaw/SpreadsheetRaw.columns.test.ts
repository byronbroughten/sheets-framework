import { describe, expect, it } from "vitest";

import {
  buildGridRows,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { SpreadsheetRaw } from "./SpreadsheetRaw";
import {
  colIdRowIndex,
  formulaCell,
  gridRanges,
  lightGreen,
  tableHeaderRowIndex,
  tableId111,
  threeByThreeAround,
  topDataRowIndex,
} from "./spreadsheetRawTestSupport";

describe("ColumnRaw.updateAllCells", () => {
  function stubFilledSheet() {
    return stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({
            0: ["c:lse:aaa", "c:lse:bbb"],
            4: ["r:lse:1", "old"],
            5: ["r:lse:2", "old"],
            6: ["r:lse:3", "old"],
          }),
          table: { endRowIndex: 7 },
        },
      ],
    });
  }
  function fetchedColumn() {
    const raw = SpreadsheetRaw.init();
    raw.sheetMeta(111).primary.column(1).gatherFetchFull();
    raw.fetchAllGathered();
    return raw;
  }

  it("fills every data row of the column, leaving the rows above it and the column beside it", () => {
    const { grid } = stubFilledSheet();

    const raw = fetchedColumn();
    raw.table(tableId111).column(1).updateAllCells({ value: "new" });
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).values()).toEqual([
      ["c:lse:aaa", "c:lse:bbb"],
      [null, null],
      [null, null],
      [null, null],
      ["r:lse:1", "new"],
      ["r:lse:2", "new"],
      ["r:lse:3", "new"],
    ]);
  });

  it("mirrors the fill into row state, so a read before the flush sees it", () => {
    stubFilledSheet();

    const raw = fetchedColumn();
    raw.table(tableId111).column(1).updateAllCells({ value: "new" });

    expect(raw.table(tableId111).column(1).valueArrOrEmpty).toEqual([
      "new",
      "new",
      "new",
    ]);
  });

  it("lets a per-cell write win over the fill", () => {
    const { grid } = stubFilledSheet();

    const raw = fetchedColumn();
    raw.table(tableId111).column(1).updateAllCells({ value: "filled" });
    raw.table(tableId111).row(1).cell(1).updateValue("overridden");
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).values(gridRanges.columnOneData)).toEqual([
      ["filled"],
      ["overridden"],
      ["filled"],
    ]);
  });

  it("fills a background colour alongside the value", () => {
    const { grid } = stubFilledSheet();

    const raw = fetchedColumn();
    raw
      .table(tableId111)
      .column(1)
      .updateAllCells({ value: "new", backgroundColor: lightGreen });
    raw.batchUpdateGSheets();

    const filled = { value: "new", backgroundColor: lightGreen };
    expect(grid.sheet(111).rows(gridRanges.columnOneData)).toEqual([
      [filled],
      [filled],
      [filled],
    ]);
  });

  it("leaves a row appended after the fill alone, since the fill's bound is snapshotted", () => {
    const { grid } = stubFilledSheet();

    const raw = fetchedColumn();
    raw.table(tableId111).column(1).updateAllCells({ value: "filled" });
    raw.table(tableId111).appendDataRow();
    raw.batchUpdateGSheets();

    expect(
      grid.sheet(111).values({ ...gridRanges.columnOneData, endRowIndex: 8 }),
    ).toEqual([["filled"], ["filled"], ["filled"], [null]]);
  });
});

describe("ColumnRaw.updateActiveCells", () => {
  function stubSelectionSheet() {
    return stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({
            0: ["c:lse:aaa", "c:lse:bbb"],
            4: ["r:lse:1", "old"],
            5: ["r:lse:2", "old"],
            6: ["r:lse:3", "old"],
            7: ["r:lse:4", "old"],
            8: ["r:lse:5", "old"],
          }),
          table: { endRowIndex: 9 },
        },
      ],
    });
  }
  function fetchedSelectionSheet() {
    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheetMeta(111).colIdRow.gatherFetchFull();
    raw.table(tableId111).column(1).gatherFetchFull();
    raw.fetchAllGathered();
    return raw;
  }
  it("fills every active row of a column whose active rows are all contiguous", () => {
    const { grid } = stubSelectionSheet();

    const raw = fetchedSelectionSheet();
    raw.table(tableId111).column(1).updateActiveCells({ value: "new" });
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).values(gridRanges.columnOneData)).toEqual([
      ["new"],
      ["new"],
      ["new"],
      ["new"],
      ["new"],
    ]);
  });

  it("fills only the active rows when they fall in separate runs", () => {
    const { grid } = stubSelectionSheet();

    const raw = fetchedSelectionSheet();
    raw.table(tableId111).removeRowsExcept(0, 1, 4);
    raw.table(tableId111).column(1).updateActiveCells({ value: "new" });
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).values(gridRanges.columnOneData)).toEqual([
      ["new"],
      ["new"],
      ["old"],
      ["old"],
      ["new"],
    ]);
  });

  it("writes value and background colour together", () => {
    const { grid } = stubSelectionSheet();

    const raw = fetchedSelectionSheet();
    raw.table(tableId111).removeRowsExcept(0);
    raw
      .table(tableId111)
      .column(1)
      .updateActiveCells({ value: "new", backgroundColor: lightGreen });
    raw.batchUpdateGSheets();

    expect(
      grid.sheet(111).rows({ ...gridRanges.columnOneData, endRowIndex: 6 }),
    ).toEqual([[{ value: "new", backgroundColor: lightGreen }], ["old"]]);
  });

  it("leaves values alone when only a background colour is written", () => {
    const { grid } = stubSelectionSheet();

    const raw = fetchedSelectionSheet();
    raw.table(tableId111).removeRowsExcept(0);
    raw
      .table(tableId111)
      .column(1)
      .updateActiveCells({ backgroundColor: lightGreen });
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).cell(4, 1)).toEqual({
      value: "old",
      backgroundColor: lightGreen,
    });
    expect(raw.table(tableId111).column(1).valueArrOrEmpty).toEqual(["old"]);
  });

  it("sends no batch update when no row is active", () => {
    const { batchUpdateCount } = stubSelectionSheet();

    const raw = fetchedSelectionSheet();
    raw.table(tableId111).removeRowsExcept();
    raw.table(tableId111).column(1).updateActiveCells({ value: "new" });
    raw.batchUpdateGSheets();

    expect(batchUpdateCount()).toBe(0);
  });

  it("mirrors the write into row state, so a read before the flush sees it", () => {
    stubSelectionSheet();

    const raw = fetchedSelectionSheet();
    raw.table(tableId111).removeRowsExcept(0, 4);
    raw.table(tableId111).column(1).updateActiveCells({ value: "new" });

    expect(raw.table(tableId111).column(1).valueArrOrEmpty).toEqual([
      "new",
      "new",
    ]);
  });
});

describe("ColumnRaw.updateAllFormulas", () => {
  function stubFilledSheet() {
    return stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({
            0: ["c:lse:aaa", "c:lse:bbb"],
            4: ["r:lse:1", "old"],
            5: ["r:lse:2", "old"],
            6: ["r:lse:3", "old"],
          }),
          table: { endRowIndex: 7 },
        },
      ],
    });
  }

  function fetchedColumn() {
    const raw = SpreadsheetRaw.init();
    raw.sheetMeta(111).primary.column(1).gatherFetchFull();
    raw.fetchAllGathered();
    return raw;
  }

  it("writes the formula into every data row of the column", () => {
    const { grid } = stubFilledSheet();

    const raw = fetchedColumn();
    raw.table(tableId111).column(1).updateAllFormulas("=2+1");
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).rows(gridRanges.columnOneData)).toEqual([
      [formulaCell("=2+1")],
      [formulaCell("=2+1")],
      [formulaCell("=2+1")],
    ]);
  });

  it("leaves row state holding the fetched values, since a formula fill isn't mirrored into it", () => {
    stubFilledSheet();

    const raw = fetchedColumn();
    raw.table(tableId111).column(1).updateAllFormulas("=2+1");
    raw.batchUpdateGSheets();

    expect(raw.table(tableId111).column(1).valueArrOrEmpty).toEqual([
      "old",
      "old",
      "old",
    ]);
  });

  it("writes a formula holding commas and quotes into one cell", () => {
    const { grid } = stubFilledSheet();
    const formula = '=FILTER(A:A,A:A<>"")';

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).row(0).cell(1).updateFormula(formula);
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).rows(threeByThreeAround(4, 1))).toEqual([
      [null, null, null],
      ["r:lse:1", formulaCell(formula), null],
      ["r:lse:2", "old", null],
    ]);
  });

  it("writes a pretty-printed formula into one cell", () => {
    const { grid } = stubFilledSheet();
    const formula = "=2+SINGLE(\ntest[Number])";

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).row(0).cell(1).updateFormula(formula);
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).rows(threeByThreeAround(4, 1))).toEqual([
      [null, null, null],
      ["r:lse:1", formulaCell(formula), null],
      ["r:lse:2", "old", null],
    ]);
  });

  it("writes a formula fill and a colour fill on one column in the same batch update", () => {
    const { batchUpdateCount, grid } = stubFilledSheet();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).column(1).updateAllFormulas("=2+1");
    raw
      .table(tableId111)
      .column(1)
      .updateAllCells({ backgroundColor: lightGreen });
    raw.batchUpdateGSheets();

    expect(batchUpdateCount()).toBe(1);
    const filled = { ...formulaCell("=2+1"), backgroundColor: lightGreen };
    expect(grid.sheet(111).rows(gridRanges.columnOneData)).toEqual([
      [filled],
      [filled],
      [filled],
    ]);
  });

  it("lets a formula written after a value on the same cell win", () => {
    const { grid } = stubFilledSheet();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    const cell = raw.table(tableId111).row(0).cell(1);
    cell.updateValue("new");
    cell.updateFormula("=2+1");
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).cell(4, 1)).toEqual(formulaCell("=2+1"));
  });

  it("lets a per-cell formula win over the formula fill", () => {
    const { grid } = stubFilledSheet();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).column(1).updateAllFormulas("=2+1");
    raw.table(tableId111).row(1).cell(1).updateFormula("=9");
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).rows(gridRanges.columnOneData)).toEqual([
      [formulaCell("=2+1")],
      [formulaCell("=9")],
      [formulaCell("=2+1")],
    ]);
  });
});

describe("SpreadsheetRaw.findReplace", () => {
  function stubFilledSheet() {
    return stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({
            [colIdRowIndex]: ["c:lse:aaa", "c:lse:bbb", "c:lse:ccc"],
            [tableHeaderRowIndex]: ["ID", "Currency", "Currency"],
            [topDataRowIndex]: ["r:lse:1", "Currency", "Currency"],
            [topDataRowIndex + 1]: ["r:lse:2", "Caretaking", "Currency"],
            [topDataRowIndex + 2]: ["r:lse:3", "Currency", "Currency"],
          }),
          table: { endRowIndex: topDataRowIndex + 3 },
        },
        {
          sheetId: 222,
          title: "Other",
          rows: [["Currency", formulaCell('="Currency"')]],
        },
      ],
    });
  }
  function fetchedColumn() {
    const raw = SpreadsheetRaw.init();
    raw.sheetMeta(111).primary.column(1).gatherFetchFull();
    raw.fetchAllGathered();
    return raw;
  }

  it("replaces within that column's data rows only", () => {
    const { grid } = stubFilledSheet();

    const raw = fetchedColumn();
    raw.table(tableId111).column(1).findReplace({
      find: "Currency",
      replacement: "Total",
      matchEntireCell: true,
    });
    raw.batchUpdateGSheets();

    expect(
      grid.sheet(111).values({ startRowIndex: tableHeaderRowIndex }),
    ).toEqual([
      ["ID", "Currency", "Currency"],
      ["r:lse:1", "Total", "Currency"],
      ["r:lse:2", "Caretaking", "Currency"],
      ["r:lse:3", "Total", "Currency"],
    ]);
  });

  it("replaces across the Table's body only, leaving its head rows and other sheets", () => {
    const { grid } = stubFilledSheet();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw
      .table(tableId111)
      .findReplace({ find: "Currency", replacement: "Total" });
    raw.batchUpdateGSheets();

    expect(
      grid.sheet(111).values({ startRowIndex: tableHeaderRowIndex }),
    ).toEqual([
      ["ID", "Currency", "Currency"],
      ["r:lse:1", "Total", "Total"],
      ["r:lse:2", "Caretaking", "Total"],
      ["r:lse:3", "Total", "Total"],
    ]);
    expect(grid.sheet(222).values()).toEqual([["Currency", '="Currency"']]);
  });

  it("replaces within a Table's body where a same-flush growth above has pushed it, and nowhere else", () => {
    const lowerHeaderRowIndex = topDataRowIndex + 2;
    const { grid } = stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({
            [tableHeaderRowIndex]: ["ID", "Name"],
            [topDataRowIndex]: ["t1", "x"],
            [lowerHeaderRowIndex]: ["ID", "x"],
            [lowerHeaderRowIndex + 1]: ["b1", "x"],
            [lowerHeaderRowIndex + 2]: ["b2", "x"],
            [lowerHeaderRowIndex + 4]: ["below", "x"],
          }),
          tables: [
            {
              tableId: "top",
              name: "Top",
              endColumnIndex: 2,
              endRowIndex: topDataRowIndex + 1,
            },
            {
              tableId: "lower",
              name: "Lower",
              startRowIndex: lowerHeaderRowIndex,
              endColumnIndex: 2,
              endRowIndex: lowerHeaderRowIndex + 3,
            },
          ],
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table("lower").findReplace({ find: "x", replacement: "z" });
    raw.table("top").appendDataRow().updateValue(1, "x");
    raw.batchUpdateGSheets();

    expect(
      grid.sheet(111).values({ startRowIndex: tableHeaderRowIndex }),
    ).toEqual([
      ["ID", "Name"],
      ["t1", "x"],
      [null, "x"],
      [null, null],
      ["ID", "x"],
      ["b1", "z"],
      ["b2", "z"],
      [null, null],
      ["below", "x"],
    ]);
  });

  it("replaces across every sheet, formulas included when asked", () => {
    const { grid } = stubFilledSheet();

    const raw = SpreadsheetRaw.init();
    raw.findReplace({
      find: "Currency",
      replacement: "Total",
      scope: { allSheets: true },
      includeFormulas: true,
    });
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).values({ startRowIndex: topDataRowIndex })).toEqual([
      ["r:lse:1", "Total", "Total"],
      ["r:lse:2", "Caretaking", "Total"],
      ["r:lse:3", "Total", "Total"],
    ]);
    expect(grid.sheet(222).rows()).toEqual([
      ["Total", formulaCell('="Total"')],
    ]);
  });

  it("replaces text the same flush writes, since it runs after the per-cell writes", () => {
    const { grid } = stubFilledSheet();

    const raw = fetchedColumn();
    raw.findReplace({
      find: "Currency",
      replacement: "Total",
      scope: { allSheets: true },
    });
    raw.table(tableId111).row(1).cell(1).updateValue("Currency");
    raw.table(tableId111).row(2).delete();
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).values(gridRanges.columnOneData)).toEqual([
      ["Total"],
      ["Total"],
    ]);
  });

  it("leaves fetched values readable until the flush actually sends", () => {
    stubFilledSheet();

    const raw = fetchedColumn();
    raw
      .table(tableId111)
      .column(1)
      .findReplace({ find: "Currency", replacement: "Total" });

    expect(raw.table(tableId111).column(1).valueArrOrEmpty).toEqual([
      "Currency",
      "Caretaking",
      "Currency",
    ]);
  });

  it("makes a read after the flush throw rather than return a pre-replace value", () => {
    stubFilledSheet();

    const raw = fetchedColumn();
    raw
      .table(tableId111)
      .column(1)
      .findReplace({ find: "Currency", replacement: "Total" });
    raw.batchUpdateGSheets();

    expect(() =>
      raw.table(tableId111).row(0).cell(1).valueOrEmpty(),
    ).toThrowError(/went stale when a findReplace was sent/);
  });

  it("leaves fetched values alone when no findReplace was queued", () => {
    stubFilledSheet();

    const raw = fetchedColumn();
    raw.table(tableId111).row(0).cell(1).updateValue("Total");
    raw.batchUpdateGSheets();

    expect(raw.table(tableId111).column(1).valueArrOrEmpty).toEqual([
      "Total",
      "Caretaking",
      "Currency",
    ]);
  });

  it("makes the values readable again after a re-fetch", () => {
    stubFilledSheet();

    const raw = fetchedColumn();
    raw
      .table(tableId111)
      .column(1)
      .findReplace({ find: "Currency", replacement: "Total" });
    raw.batchUpdateGSheets();
    raw.table(tableId111).column(1).gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.table(tableId111).column(1).valueArrOrEmpty).toEqual([
      "Total",
      "Caretaking",
      "Total",
    ]);
  });

  it("discards a queued replace alongside every other change, sending no batch update", () => {
    const { batchUpdateCount } = stubFilledSheet();

    const raw = SpreadsheetRaw.init();
    raw.findReplace({
      find: "Currency",
      replacement: "Total",
      scope: { allSheets: true },
    });
    raw.discardQueuedChanges();
    raw.batchUpdateGSheets();

    expect(batchUpdateCount()).toBe(0);
  });
});
