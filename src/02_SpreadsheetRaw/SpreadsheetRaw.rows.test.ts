import { describe, expect, it } from "vitest";

import {
  type AddTableOperation,
  type BoundedGridRange,
} from "../00_Source/RawSource/RawSource";
import { SheetIndex } from "../00_Source/RawSource/SheetIndex";
import {
  buildGridRows,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { SpreadsheetRaw } from "./SpreadsheetRaw";
import {
  formulaCell,
  scratchGid,
  startTableColIndex,
  tableHeaderRowIndex,
  topDataRowIndex,
} from "./spreadsheetRawTestSupport";

describe("SpreadsheetRaw add sheet and add Table", () => {
  const addSheetProps = {
    sheetId: 555,
    title: "Spreadsheet Config",
    rowCount: 20,
    columnCount: 6,
  };
  const addTableProps: Omit<AddTableOperation, "kind"> = {
    name: "spreadsheetConfig",
    range: {
      sheetId: 555,
      startRowIndex: SheetIndex.row(2),
      endRowIndex: SheetIndex.row(5),
      startColumnIndex: SheetIndex.col(1),
      endColumnIndex: SheetIndex.col(3),
    },
    columnProperties: [
      { columnIndex: 1, columnName: "Name", columnType: "TEXT" },
    ],
  };

  const addedTable = {
    tableId: "spreadsheetConfig",
    name: "spreadsheetConfig",
    range: {
      startRowIndex: 2,
      endRowIndex: 5,
      startColumnIndex: 1,
      endColumnIndex: 3,
    },
    columnProperties: [
      {},
      { columnIndex: 1, columnName: "Name", columnType: "TEXT" },
    ],
  };

  it("adds the tab and its Table in the same batch update that renames another tab, whatever the gather order", () => {
    const { batchUpdateCount, grid } = stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).updateTitle("Renamed");
    raw.gatherAddTableOperation(addTableProps);
    raw.gatherAddSheetOperation(addSheetProps);
    raw.batchUpdateGSheets();

    expect(batchUpdateCount()).toBe(1);
    expect(grid.sheetTitles()).toEqual(["Renamed", "Spreadsheet Config"]);
    const added = grid.sheet(555);
    expect([added.rowCount, added.columnCount]).toEqual([20, 6]);
    expect(added.tables).toEqual([addedTable]);
  });

  it("puts one operation on the spreadsheet's write queue per queue method", () => {
    stubSheetsService();

    const raw = SpreadsheetRaw.init();
    raw.gatherAddSheetOperation(addSheetProps);
    raw.gatherAddTableOperation(addTableProps);

    expect(raw.writeOperations.addSheet).toEqual([
      { kind: "addSheet", ...addSheetProps },
    ]);
    expect(raw.writeOperations.addTable).toEqual([
      { kind: "addTable", ...addTableProps },
    ]);
  });

  it("empties both lists on a flush", () => {
    stubSheetsService();

    const raw = SpreadsheetRaw.init();
    raw.gatherAddSheetOperation(addSheetProps);
    raw.gatherAddTableOperation(addTableProps);
    raw.batchUpdateGSheets();

    expect(raw.writeOperations.addSheet).toEqual([]);
    expect(raw.writeOperations.addTable).toEqual([]);
  });

  const seededCell = {
    sheetId: 555,
    rowIndex: SheetIndex.row(3),
    colIndex: SheetIndex.col(1),
    value: "Example",
  };

  const checkboxRange: BoundedGridRange = {
    sheetId: 555,
    startRowIndex: SheetIndex.row(3),
    endRowIndex: SheetIndex.row(4),
    startColumnIndex: SheetIndex.col(1),
    endColumnIndex: SheetIndex.col(2),
  };

  it("refuses a seeded value for a GID with no add-sheet queued, naming the GID, and queues nothing", () => {
    stubSheetsService();

    const raw = SpreadsheetRaw.init();

    expect(() => raw.gatherAddedSheetFillCellOperation(seededCell)).toThrow(
      "Added-sheet cell write refused: no addSheet for GID 555 is queued in this flush.",
    );
    expect(raw.writeOperations.fillCell).toEqual([]);
  });

  it("refuses a seeded value after the create flush has sent that GID's add-sheet", () => {
    stubSheetsService();

    const raw = SpreadsheetRaw.init();
    raw.gatherAddSheetOperation(addSheetProps);
    raw.batchUpdateGSheets();

    expect(() => raw.gatherAddedSheetFillCellOperation(seededCell)).toThrow(
      "no addSheet for GID 555",
    );
  });

  const { value: _value, ...seededPosition } = seededCell;

  it("seeds a formula on an added tab", () => {
    const { grid } = stubSheetsService();

    const raw = SpreadsheetRaw.init();
    raw.gatherAddSheetOperation(addSheetProps);
    raw.gatherAddedSheetFillCellOperation({
      ...seededPosition,
      formula: "=ROW()",
    });
    raw.batchUpdateGSheets();

    expect(grid.sheet(555).cell(3, 1)).toEqual(formulaCell("=ROW()"));
  });

  it("refuses a seeded formula that doesn't start with =", () => {
    stubSheetsService();

    const raw = SpreadsheetRaw.init();
    raw.gatherAddSheetOperation(addSheetProps);

    expect(() =>
      raw.gatherAddedSheetFillCellOperation({
        ...seededPosition,
        formula: "ROW()",
      }),
    ).toThrow('Formula must start with "="');
  });

  it("seeds a value inside the added Table in the same batch update that adds the tab", () => {
    const { batchUpdateCount, grid } = stubSheetsService();

    const raw = SpreadsheetRaw.init();
    raw.gatherAddSheetOperation(addSheetProps);
    raw.gatherAddTableOperation(addTableProps);
    raw.gatherAddedSheetFillCellOperation(seededCell);
    raw.batchUpdateGSheets();

    expect(batchUpdateCount()).toBe(1);
    expect(grid.sheet(555).cell(3, 1)).toBe("Example");
    expect(grid.sheet(555).tables).toEqual([addedTable]);
  });

  it("refuses a checkbox validation for a GID with no add-sheet queued, naming the GID, and queues nothing", () => {
    stubSheetsService();

    const raw = SpreadsheetRaw.init();

    expect(() =>
      raw.gatherAddedSheetCheckboxValidationOperation(checkboxRange),
    ).toThrowError(
      "Added-sheet checkbox validation refused: no addSheet for GID 555 is queued in this flush.",
    );
    expect(raw.writeOperations.addCheckboxValidation).toEqual([]);
  });

  it("makes a seeded cell on an added tab a checkbox holding its seeded value, in the same batch update", () => {
    const { batchUpdateCount, grid } = stubSheetsService();

    const raw = SpreadsheetRaw.init();
    raw.gatherAddSheetOperation(addSheetProps);
    raw.gatherAddedSheetCheckboxValidationOperation(checkboxRange);
    raw.gatherAddTableOperation(addTableProps);
    raw.gatherAddedSheetFillCellOperation({ ...seededCell, value: false });
    raw.batchUpdateGSheets();

    expect(batchUpdateCount()).toBe(1);
    expect(grid.sheet(555).cell(3, 1)).toEqual({
      value: false,
      dataValidationConditionType: "BOOLEAN",
    });
    expect(raw.writeOperations.addCheckboxValidation).toEqual([]);
  });

  it("drops a queued seeded value on discardQueuedChanges, sending no batch update", () => {
    const { batchUpdateCount } = stubSheetsService();

    const raw = SpreadsheetRaw.init();
    raw.gatherAddSheetOperation(addSheetProps);
    raw.gatherAddedSheetFillCellOperation(seededCell);
    raw.discardQueuedChanges();
    raw.batchUpdateGSheets();

    expect(raw.writeOperations.fillCell).toEqual([]);
    expect(batchUpdateCount()).toBe(0);
  });
});

describe("RowRaw.delete", () => {
  function stubSheetWithDataRows(dataRowCount: number) {
    return stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows(
            Object.fromEntries(
              Array.from({ length: dataRowCount }, (_, offset) => [
                topDataRowIndex + offset,
                [`r${offset + 1}`],
              ]),
            ),
          ),
          table: { endRowIndex: topDataRowIndex + dataRowCount },
        },
      ],
    });
  }

  it("refuses to delete the only data row, since a new row copies its formulas from the rows already there", () => {
    stubSheetWithDataRows(1);

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();

    expect(() => raw.sheet(111).topRow.delete()).toThrowError(
      /last data row.*may never be left with none/,
    );
  });

  it("refuses the delete that would take the last of several to zero", () => {
    stubSheetWithDataRows(3);

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).row(0).delete();
    raw.sheet(111).row(1).delete();

    expect(() => raw.sheet(111).row(2).delete()).toThrowError(
      /last data row.*may never be left with none/,
    );
  });

  it("still deletes a row when other data rows survive it", () => {
    const { grid } = stubSheetWithDataRows(2);

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).row(1).delete();
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).values({ startRowIndex: topDataRowIndex })).toEqual([
      ["r1"],
    ]);
  });

  it("counts a row appended and then deleted as neither, since the two cancel before the flush", () => {
    stubSheetWithDataRows(1);

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).appendDataRow().delete();

    expect(() => raw.sheet(111).topRow.delete()).toThrowError(/last data row/);
  });

  it("does not change another sheet's data-row count when a row delete is queued", () => {
    stubSheetsService({
      sheets: [
        { sheetId: 111, title: "Records", table: { endRowIndex: 11 } },
        { sheetId: 222, title: "Entries", table: { endRowIndex: 6 } },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    const unitsBefore = raw.sheet(222).dataRowCountAfterFlush;

    raw.sheet(111).row(1).delete();

    expect(raw.sheet(222).dataRowCountAfterFlush).toBe(unitsBefore);
    expect(raw.sheet(111).dataRowCountAfterFlush).toBe(6);
  });

  it("spans a column inserted at the Table end in the same flush, so the delete never covers part of the Table", () => {
    const { grid } = stubSheetWithDataRows(2);

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheetMeta(111).insertColumnAtEnd({ columnId: "c:x:new", header: "New" });
    raw.sheet(111).row(0).delete();
    raw.batchUpdateGSheets();

    expect(
      grid.sheet(111).values({
        startRowIndex: topDataRowIndex,
        endRowIndex: topDataRowIndex + 1,
      }),
    ).toEqual([["r2", null]]);
  });

  describe("beside a neighbouring Table", () => {
    const leftStart = startTableColIndex;
    const rightStart = startTableColIndex + 3;
    const bothTablesBand = {
      startRowIndex: topDataRowIndex,
      endRowIndex: topDataRowIndex + 3,
      startColumnIndex: leftStart,
      endColumnIndex: rightStart + 2,
    };
    function stubSideBySideTables() {
      return stubSheetsService({
        sheets: [
          {
            sheetId: scratchGid,
            title: "Byron's Scratch Sheet",
            rows: buildGridRows({
              [tableHeaderRowIndex]: ["ID", "Name", "", "Code", "Qty"],
              [topDataRowIndex]: ["r1", "a", "loose", "c1", 5],
              [topDataRowIndex + 1]: ["r2", "b", "", "c2", 6],
              [topDataRowIndex + 2]: ["r3", "c", "", "c3", 7],
            }),
            tables: [
              {
                tableId: "left",
                startColumnIndex: leftStart,
                endColumnIndex: leftStart + 2,
                endRowIndex: topDataRowIndex + 3,
              },
              {
                tableId: "right",
                startColumnIndex: rightStart,
                endColumnIndex: rightStart + 2,
                endRowIndex: topDataRowIndex + 3,
              },
            ],
          },
        ],
      });
    }

    it("shifts up only the deleting Table's columns, leaving the neighbour and the cells between untouched", () => {
      const { grid } = stubSideBySideTables();

      const raw = SpreadsheetRaw.init();
      raw.fetchAllSheetProperties();
      raw.table("left").row(0).delete();
      raw.batchUpdateGSheets();

      expect(grid.sheet(scratchGid).values(bothTablesBand)).toEqual([
        ["r2", "b", "loose", "c1", 5],
        ["r3", "c", "", "c2", 6],
        [null, null, "", "c3", 7],
      ]);
    });

    it("lands several deletes in one Table on the rows they named", () => {
      const { grid } = stubSideBySideTables();

      const raw = SpreadsheetRaw.init();
      raw.fetchAllSheetProperties();
      raw.table("right").row(0).delete();
      raw.table("right").row(2).delete();
      raw.batchUpdateGSheets();

      expect(grid.sheet(scratchGid).values(bothTablesBand)).toEqual([
        ["r1", "a", "loose", "c2", 6],
        ["r2", "b", "", null, null],
        ["r3", "c", "", null, null],
      ]);
    });

    it("flags the delete on its own Table only, keyed by Table and row index", () => {
      stubSideBySideTables();

      const raw = SpreadsheetRaw.init();
      raw.fetchAllSheetProperties();
      raw.table("left").row(1).delete();

      expect(raw.table("left").row(1).isQueuedForDelete).toBe(true);
      expect(raw.table("right").row(1).isQueuedForDelete).toBe(false);
      expect(raw.table("left").dataRowCountAfterFlush).toBe(2);
      expect(raw.table("right").dataRowCountAfterFlush).toBe(3);
    });
  });
});

describe("RowRaw.rowIsActive", () => {
  it("makes an appended row active and grows the table end before the flush", () => {
    stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    const countBefore = raw.sheet(111).dataRowCount;

    const row = raw.sheet(111).appendDataRow();

    expect(row.rowIsActive()).toBe(true);
    expect(raw.sheet(111).dataRowCount).toBe(countBefore + 1);
  });

  it("drops a removed row from the working view before the flush, leaving table indexes unmoved", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({ 4: ["kept"], 5: ["deleted"] }),
          table: { endRowIndex: 11 },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).row(1).gatherFetchFull();
    raw.fetchAllGathered();
    expect(raw.sheet(111).row(1).rowIsActive()).toBe(true);
    const countBefore = raw.sheet(111).dataRowCount;

    raw.sheet(111).row(1).delete();

    expect(raw.sheet(111).row(1).rowIsActive()).toBe(false);
    expect(raw.sheet(111).dataRowCount).toBe(countBefore);
    expect(raw.sheet(111).dataRowCountAfterFlush).toBe(6);
  });

  it("keeps pre-flush row indexes after a flushed delete, so the removed row stays inactive at its old index", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({ 4: ["kept"], 5: ["deleted"] }),
          table: { endRowIndex: 11 },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheet(111).row(1).gatherFetchFull();
    raw.fetchAllGathered();
    raw.sheet(111).row(1).delete();
    raw.batchUpdateGSheets();

    expect(raw.sheet(111).row(1).rowIsActive()).toBe(false);
    expect(raw.sheet(111).rowIndexesAreStale).toBe(true);
    expect(() => raw.sheet(111).dataRowCount).toThrow(/Row indexes are stale/);
  });
});

describe("TableRaw.removeRowsExcept", () => {
  function stubPrunableSheet() {
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
  function fetchedPrunableSheet() {
    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheetMeta(111).colIdRow.gatherFetchFull();
    raw.sheet(111).column(1).gatherFetchFull();
    raw.fetchAllGathered();
    return raw;
  }

  it("drops every data row that was not kept", () => {
    stubPrunableSheet();

    const raw = fetchedPrunableSheet();
    raw.sheet(111).removeRowsExcept(1);

    expect(raw.sheet(111).rowIndexesActive).toEqual([1]);
  });

  it("keeps the uniform rows, so a column still resolves by its id afterwards", () => {
    stubPrunableSheet();

    const raw = fetchedPrunableSheet();
    raw.sheet(111).removeRowsExcept(1);

    expect(raw.sheetMeta(111).columnByActiveId("c:lse:bbb").colIndex).toBe(1);
  });

  it("makes a whole-column fill throw, so it can't overwrite the excluded rows", () => {
    stubPrunableSheet();

    const raw = fetchedPrunableSheet();
    raw.sheet(111).removeRowsExcept(1);

    expect(() =>
      raw.sheet(111).column(1).updateAllCells({ value: "new" }),
    ).toThrowError(/pruned to a selection/);
  });
});

describe("TableRaw.dataRowCount", () => {
  function fetchedSheet(endRowIndex: number) {
    stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          table: { endRowIndex },
        },
      ],
    });
    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    return raw.sheet(111);
  }

  it("throws when the exclusive end row is the first data row", () => {
    expect(() => fetchedSheet(topDataRowIndex).dataRowCount).toThrow(
      /Records.*at least one data row/,
    );
  });

  it("accepts a Table whose exclusive end is one past the first data row", () => {
    expect(fetchedSheet(topDataRowIndex + 1).dataRowCount).toBe(1);
  });
});
