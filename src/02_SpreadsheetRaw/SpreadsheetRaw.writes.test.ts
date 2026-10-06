import { describe, expect, it } from "vitest";

import { googleRawRequest } from "../00_Source/GoogleSheets/GoogleSheetsAPI";
import {
  buildGridRows,
  type FakeCell,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { SpreadsheetRaw } from "./SpreadsheetRaw";
import {
  firstTableEndRowIndex,
  gridRanges,
  lightGreen,
  startTableColIndex,
  tableHeaderRowIndex,
  tableId111,
  tableId222,
  threeByThreeAround,
  topDataRowIndex,
} from "./spreadsheetRawTestSupport";

describe("SpreadsheetRaw.batchUpdateGSheets", () => {
  it("sends nothing for a write queued before the sheet's Table was fetched, once discarded", () => {
    const { batchUpdateCount } = stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records" }],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheetMeta(111).primary.requestSortGSheet({
      colIdxToSortBy: 0,
      sortOrder: "ASCENDING",
    });
    raw.discardQueuedChanges();
    raw.batchUpdateGSheets();

    expect(batchUpdateCount()).toBe(0);
  });

  it("sends no batch update when there is nothing to save", () => {
    const { batchUpdateCount } = stubSheetsService();

    SpreadsheetRaw.init().batchUpdateGSheets();

    expect(batchUpdateCount()).toBe(0);
  });

  it("grows the Table by every appended row rather than by one", () => {
    const { grid } = stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          table: { endRowIndex: 11, endColumnIndex: 2 },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).appendDataRow();
    raw.table(tableId111).appendDataRow();
    raw.table(tableId111).appendDataRow();
    raw.batchUpdateGSheets();

    expect(firstTableEndRowIndex(grid, 111)).toBe(14);
  });

  it("grows each sheet's Table by the rows appended to that sheet, however the appends interleave", () => {
    const { grid } = stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          table: { endRowIndex: 11, endColumnIndex: 2 },
        },
        {
          sheetId: 222,
          title: "Entries",
          table: { endRowIndex: 6, endColumnIndex: 2 },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).appendDataRow();
    raw.table(tableId222).appendDataRow();
    raw.table(tableId111).appendDataRow();
    raw.batchUpdateGSheets();

    expect(firstTableEndRowIndex(grid, 111)).toBe(13);
    expect(firstTableEndRowIndex(grid, 222)).toBe(7);
  });

  it("still appends a row queued after a deletion on the same sheet, since row indexes only shift once the deletes are sent", () => {
    const { grid } = stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({ 4: ["kept"], 5: ["deleted"], 6: ["later"] }),
          table: { endRowIndex: 11 },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).row(1).delete();
    raw.table(tableId111).appendDataRow();

    expect(() => raw.batchUpdateGSheets()).not.toThrow();
    expect(grid.sheet(111).values({ startRowIndex: 4 })).toEqual([
      ["kept"],
      ["later"],
    ]);
    expect(firstTableEndRowIndex(grid, 111)).toBe(11);
    expect(raw.table(tableId111).rowIndexesAreStale).toBe(true);
  });

  describe("Table growth", () => {
    const dateColIndex = startTableColIndex + 3;
    function stubGrowingTable() {
      return stubSheetsService({
        sheets: [
          {
            sheetId: 111,
            title: "Records",
            rows: buildGridRows({
              [tableHeaderRowIndex]: ["ID", "Amount", "Done", "When"],
              [topDataRowIndex]: [
                "r1",
                {
                  value: 5,
                  backgroundColor: lightGreen,
                  numberFormatType: "CURRENCY",
                },
                { value: false, dataValidationConditionType: "BOOLEAN" },
                { value: 1, dataValidationConditionType: "DATE_IS_VALID" },
              ],
            }),
            table: {
              endRowIndex: topDataRowIndex + 1,
              columnTypes: { [dateColIndex]: "DATE" },
            },
          },
        ],
      });
    }

    it("extends the grid by exactly the rows needed, widens the Table over them, carries format and untyped validation down, and fills them", () => {
      const { grid } = stubGrowingTable();
      expect(grid.sheet(111).rowCount).toBe(topDataRowIndex + 1);

      const raw = SpreadsheetRaw.init();
      raw.fetchAllSheetProperties();
      raw.table(tableId111).appendDataRow().updateValue(0, "r2");
      raw.table(tableId111).appendDataRow().cell(1).updateFormula("=1+1");
      raw.batchUpdateGSheets();

      const sheet = grid.sheet(111);
      expect(sheet.rowCount).toBe(topDataRowIndex + 3);
      expect(firstTableEndRowIndex(grid, 111)).toBe(topDataRowIndex + 3);
      expect(
        sheet.rows({
          startRowIndex: topDataRowIndex + 1,
          endRowIndex: topDataRowIndex + 3,
          startColumnIndex: startTableColIndex,
          endColumnIndex: startTableColIndex + 4,
        }),
      ).toEqual([
        [
          "r2",
          {
            value: null,
            backgroundColor: lightGreen,
            numberFormatType: "CURRENCY",
          },
          { value: null, dataValidationConditionType: "BOOLEAN" },
          null,
        ],
        [
          null,
          {
            value: "=1+1",
            isFormula: true,
            backgroundColor: lightGreen,
            numberFormatType: "CURRENCY",
          },
          { value: null, dataValidationConditionType: "BOOLEAN" },
          null,
        ],
      ]);
    });

    it("grows from the last body row even after a same-run re-fetch resets the Table's end", () => {
      const { grid } = stubGrowingTable();

      const raw = SpreadsheetRaw.init();
      raw.fetchAllSheetProperties();
      raw.table(tableId111).appendDataRow().updateValue(0, "r2");
      raw.table(tableId111).appendDataRow().updateValue(0, "r3");
      raw.fetchAllSheetProperties();
      raw.batchUpdateGSheets();

      expect(firstTableEndRowIndex(grid, 111)).toBe(topDataRowIndex + 3);
      expect(
        grid.sheet(111).values({
          startRowIndex: topDataRowIndex,
          endColumnIndex: startTableColIndex + 1,
        }),
      ).toEqual([["r1"], ["r2"], ["r3"]]);
    });

    it("leaves no gap when an appended row is deleted in the same flush", () => {
      const { grid } = stubGrowingTable();

      const raw = SpreadsheetRaw.init();
      raw.fetchAllSheetProperties();
      const deleted = raw
        .table(tableId111)
        .appendDataRow()
        .updateValue(0, "gone");
      raw.table(tableId111).appendDataRow().updateValue(0, "r2");
      deleted.delete();
      raw.batchUpdateGSheets();

      expect(firstTableEndRowIndex(grid, 111)).toBe(topDataRowIndex + 2);
      expect(
        grid.sheet(111).values({
          startRowIndex: topDataRowIndex,
          endColumnIndex: startTableColIndex + 1,
        }),
      ).toEqual([["r1"], ["r2"]]);
    });

    it("needs no appendDimension when the grid already reaches past the new rows", () => {
      const { grid } = stubSheetsService({
        sheets: [
          {
            sheetId: 111,
            title: "Records",
            rows: buildGridRows({ [topDataRowIndex + 5]: ["below"] }),
            table: { endRowIndex: topDataRowIndex + 1 },
          },
        ],
      });

      const raw = SpreadsheetRaw.init();
      raw.fetchAllSheetProperties();
      raw.table(tableId111).appendDataRow();
      raw.batchUpdateGSheets();

      expect(grid.sheet(111).rowCount).toBe(topDataRowIndex + 7);
      expect(grid.sheet(111).cell(topDataRowIndex + 6, 0)).toBe("below");
    });

    const leftStart = startTableColIndex;
    const rightStart = startTableColIndex + 3;
    function stubSideBySideTables() {
      return stubSheetsService({
        sheets: [
          {
            sheetId: 111,
            title: "Records",
            rows: buildGridRows({
              [tableHeaderRowIndex]: ["ID", "Name", "", "Code", "Qty"],
              [topDataRowIndex]: ["r1", "a", "loose", "c1", 5],
              [topDataRowIndex + 1]: ["r2", "b", "", "c2", 6],
            }),
            tables: [
              {
                tableId: "left",
                startColumnIndex: leftStart,
                endColumnIndex: leftStart + 2,
                endRowIndex: topDataRowIndex + 2,
              },
              {
                tableId: "right",
                startColumnIndex: rightStart,
                endColumnIndex: rightStart + 2,
                endRowIndex: topDataRowIndex + 2,
              },
            ],
          },
        ],
      });
    }
    function tableEndRowIndexes(
      grid: ReturnType<typeof stubSheetsService>["grid"],
    ): (number | undefined)[] {
      return grid.sheet(111).tables.map(({ range }) => range?.endRowIndex);
    }

    it("gives a side-by-side neighbour and the cells between no blank rows", () => {
      const { grid } = stubSideBySideTables();

      const raw = SpreadsheetRaw.init();
      raw.fetchAllSheetProperties();
      raw.table("left").appendDataRow().updateValue(0, "r3");
      raw.batchUpdateGSheets();

      expect(
        grid.sheet(111).values({ startRowIndex: topDataRowIndex }),
      ).toEqual([
        ["r1", "a", "loose", "c1", 5],
        ["r2", "b", "", "c2", 6],
        ["r3", null, null, null, null],
      ]);
      expect(tableEndRowIndexes(grid)).toEqual([
        topDataRowIndex + 3,
        topDataRowIndex + 2,
      ]);
    });

    it("grows two Tables on one sheet each by its own rows, not as one merged append", () => {
      const { grid } = stubSideBySideTables();

      const raw = SpreadsheetRaw.init();
      raw.fetchAllSheetProperties();
      raw.table("right").appendDataRow().updateValue(0, "c3");
      raw.table("left").appendDataRow().updateValue(0, "r3");
      raw.table("right").appendDataRow().updateValue(0, "c4");
      raw.batchUpdateGSheets();

      expect(tableEndRowIndexes(grid)).toEqual([
        topDataRowIndex + 3,
        topDataRowIndex + 4,
      ]);
      expect(grid.sheet(111).rowCount).toBe(topDataRowIndex + 4);
      expect(
        grid.sheet(111).values({ startRowIndex: topDataRowIndex + 2 }),
      ).toEqual([
        ["r3", null, null, "c3", null],
        [null, null, null, "c4", null],
      ]);
    });

    it("costs no round trip beyond the one batch update", () => {
      const { batchUpdateCount, getCalls, getByDataFilterCalls } =
        stubSideBySideTables();

      const raw = SpreadsheetRaw.init();
      raw.fetchAllSheetProperties();
      const fetchCount = getCalls.length + getByDataFilterCalls.length;
      raw.table("left").appendDataRow();
      raw.table("right").appendDataRow();
      raw.batchUpdateGSheets();

      expect(getCalls.length + getByDataFilterCalls.length).toBe(fetchCount);
      expect(batchUpdateCount()).toBe(1);
    });

    describe("stacked Tables", () => {
      const lowerHeaderRowIndex = topDataRowIndex + 3;
      const asideStart = startTableColIndex + 3;
      function stubStackedTables() {
        return stubSheetsService({
          sheets: [
            {
              sheetId: 111,
              title: "Records",
              rows: buildGridRows({
                [tableHeaderRowIndex]: ["ID", "Name"],
                [topDataRowIndex]: ["t1", "a"],
                [topDataRowIndex + 1]: ["t2", "b"],
                [lowerHeaderRowIndex]: ["ID", "Name", "", "Code", "Qty"],
                [lowerHeaderRowIndex + 1]: ["b1", "x", "", "c1", 5],
                [lowerHeaderRowIndex + 2]: ["b2", "y", "", "c2", 6],
              }),
              tables: [
                {
                  tableId: "top",
                  name: "Top",
                  endColumnIndex: startTableColIndex + 2,
                  endRowIndex: topDataRowIndex + 2,
                },
                {
                  tableId: "lower",
                  name: "Lower",
                  startRowIndex: lowerHeaderRowIndex,
                  endColumnIndex: startTableColIndex + 2,
                  endRowIndex: lowerHeaderRowIndex + 3,
                },
                {
                  tableId: "aside",
                  name: "Aside",
                  startRowIndex: lowerHeaderRowIndex,
                  startColumnIndex: asideStart,
                  endColumnIndex: asideStart + 2,
                  endRowIndex: lowerHeaderRowIndex + 3,
                },
              ],
            },
          ],
        });
      }
      function tableRows(
        grid: ReturnType<typeof stubSheetsService>["grid"],
      ): Record<string, [number | undefined, number | undefined]> {
        return Object.fromEntries(
          grid
            .sheet(111)
            .tables.map(({ tableId, range }) => [
              tableId,
              [range?.startRowIndex, range?.endRowIndex],
            ]),
        );
      }

      it("pushes a Table below a grown one down intact, and lands a later write in the batch on its rows after growth", () => {
        const { grid } = stubStackedTables();

        const raw = SpreadsheetRaw.init();
        raw.fetchAllSheetProperties();
        raw.table("top").appendDataRow().updateValue(0, "t3");
        raw.table("lower").row(1).cell(1).updateValue("z");
        raw.batchUpdateGSheets();

        expect(tableRows(grid)).toEqual({
          top: [tableHeaderRowIndex, topDataRowIndex + 3],
          lower: [lowerHeaderRowIndex + 1, lowerHeaderRowIndex + 4],
          aside: [lowerHeaderRowIndex, lowerHeaderRowIndex + 3],
        });
        expect(
          grid.sheet(111).values({
            startRowIndex: lowerHeaderRowIndex + 1,
            endColumnIndex: startTableColIndex + 2,
          }),
        ).toEqual([
          ["ID", "Name"],
          ["b1", "x"],
          ["b2", "z"],
        ]);
      });

      it("grows the lower Table too, filling its new row below its pushed-down rows", () => {
        const { grid } = stubStackedTables();

        const raw = SpreadsheetRaw.init();
        raw.fetchAllSheetProperties();
        raw.table("lower").appendDataRow().updateValue(0, "b3");
        raw.table("top").appendDataRow().updateValue(0, "t3");
        raw.batchUpdateGSheets();

        expect(tableRows(grid)).toMatchObject({
          top: [tableHeaderRowIndex, topDataRowIndex + 3],
          lower: [lowerHeaderRowIndex + 1, lowerHeaderRowIndex + 5],
        });
        expect(
          grid.sheet(111).values({
            startRowIndex: lowerHeaderRowIndex + 2,
            endColumnIndex: startTableColIndex + 1,
          }),
        ).toEqual([["b1"], ["b2"], ["b3"]]);
      });

      it("deletes a lower Table's row at its rows after growth", () => {
        const { grid } = stubStackedTables();

        const raw = SpreadsheetRaw.init();
        raw.fetchAllSheetProperties();
        raw.table("top").appendDataRow().updateValue(0, "t3");
        raw.table("lower").row(0).delete();
        raw.batchUpdateGSheets();

        expect(
          grid.sheet(111).values({
            startRowIndex: lowerHeaderRowIndex + 1,
            endColumnIndex: startTableColIndex + 1,
          }),
        ).toEqual([["ID"], ["b2"]]);
      });

      function flushedGrowthOfTop() {
        const { grid } = stubStackedTables();
        const raw = SpreadsheetRaw.init();
        raw.fetchAllSheetProperties();
        raw.table("top").appendDataRow().updateValue(0, "new");
        raw.batchUpdateGSheets();
        return { raw, grid };
      }

      it("lands a later flush's writes on the grown Table and on the Table it pushed down", () => {
        const { raw, grid } = flushedGrowthOfTop();

        raw.table("top").row(2).cell(1).updateValue("c");
        raw.table("lower").row(0).cell(1).updateValue("w");
        raw.table("lower").row(1).delete();
        raw.batchUpdateGSheets();

        expect(
          grid.sheet(111).values({
            startRowIndex: tableHeaderRowIndex,
            endRowIndex: lowerHeaderRowIndex + 3,
            endColumnIndex: startTableColIndex + 2,
          }),
        ).toEqual([
          ["ID", "Name"],
          ["t1", "a"],
          ["t2", "b"],
          ["new", "c"],
          [null, null],
          ["ID", "Name"],
          ["b1", "w"],
        ]);
      });

      it("grows the grown Table again in a later flush, pushing the Table below down again", () => {
        const { raw, grid } = flushedGrowthOfTop();

        raw.table("top").appendDataRow().updateValue(0, "again");
        raw.table("lower").row(1).cell(1).updateValue("z");
        raw.batchUpdateGSheets();

        expect(tableRows(grid)).toMatchObject({
          top: [tableHeaderRowIndex, topDataRowIndex + 4],
          lower: [lowerHeaderRowIndex + 2, lowerHeaderRowIndex + 5],
        });
        expect(
          grid.sheet(111).values({
            startRowIndex: topDataRowIndex + 2,
            endRowIndex: lowerHeaderRowIndex + 5,
            endColumnIndex: startTableColIndex + 2,
          }),
        ).toEqual([
          ["new", null],
          ["again", null],
          [null, null],
          ["ID", "Name"],
          ["b1", "x"],
          ["b2", "z"],
        ]);
      });

      it("flags no Table's row indexes stale after growth", () => {
        const { raw } = flushedGrowthOfTop();

        expect(raw.table("top").rowIndexesAreStale).toBe(false);
        expect(raw.table("lower").rowIndexesAreStale).toBe(false);
        expect(raw.table("aside").rowIndexesAreStale).toBe(false);
      });

      it("still refuses a row write through the sheet, which resolves to no fetched Table, after growth pushed its Tables down", () => {
        const { raw } = flushedGrowthOfTop();

        expect(() => raw.sheet(111).row(0).cell(0).updateValue("late")).toThrow(
          /sheet properties have been fetched/,
        );
      });

      it("still flags the sheet's Tables stale after a flushed row delete that follows growth, naming each", () => {
        const { raw } = flushedGrowthOfTop();

        raw.table("top").row(0).delete();
        raw.batchUpdateGSheets();

        expect(() =>
          raw.table("top").row(0).cell(0).updateValue("late"),
        ).toThrow(/Table "Top" on "Records" \(gid 111\).*refetch/);
        expect(() => raw.table("lower").row(0).delete()).toThrow(
          /Table "Lower" on "Records" \(gid 111\).*refetch/,
        );
      });

      it("rewords Google's refusal of a growth blocked by a wider Table below, naming both Tables, the sheet and the fix", () => {
        const { grid } = stubSheetsService({
          sheets: [
            {
              sheetId: 111,
              title: "Records",
              rows: buildGridRows({
                [tableHeaderRowIndex]: ["ID", "Name"],
                [topDataRowIndex]: ["t1", "a"],
                [lowerHeaderRowIndex]: ["ID", "Name", "Code"],
                [lowerHeaderRowIndex + 1]: ["b1", "x", "c1"],
              }),
              tables: [
                {
                  tableId: "top",
                  name: "Top",
                  endColumnIndex: startTableColIndex + 2,
                  endRowIndex: topDataRowIndex + 1,
                },
                {
                  tableId: "wide",
                  name: "Wide",
                  startRowIndex: lowerHeaderRowIndex,
                  endColumnIndex: startTableColIndex + 3,
                  endRowIndex: lowerHeaderRowIndex + 2,
                },
              ],
            },
          ],
        });

        const raw = SpreadsheetRaw.init();
        raw.fetchAllSheetProperties();
        raw.table("top").appendDataRow().updateValue(0, "t2");

        expect(() => raw.batchUpdateGSheets()).toThrow(
          'Growing Table "Top" on sheet "Records" would insert cells over part of Table "Wide" below it. Make the lower Table, "Wide", no wider than "Top", or move it.',
        );
        expect(tableRows(grid)).toEqual({
          top: [tableHeaderRowIndex, topDataRowIndex + 1],
          wide: [lowerHeaderRowIndex, lowerHeaderRowIndex + 2],
        });
      });
    });
  });

  const staleRowIndexes =
    'Row indexes are stale for Table "" on "Records" (gid 111): a flush has moved its rows, so it needs a refetch';

  function sheetAfterFlushedDataRowDelete(fetchKeptRow = false) {
    stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({
            4: ["kept"],
            5: ["deleted"],
            6: ["later"],
          }),
          table: {
            endRowIndex: 11,
            endColumnIndex: 3,
            columnTypes: { 0: "TEXT" },
            columnValidationValues: { 0: ["=valueConfig[Notes]"] },
            columnValidationConditionTypes: { 0: "BOOLEAN" },
          },
        },
      ],
    });
    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    if (fetchKeptRow) {
      raw.table(tableId111).row(0).gatherFetchFull();
      raw.fetchAllGathered();
    }
    raw.table(tableId111).row(1).delete();
    raw.batchUpdateGSheets();
    return raw;
  }

  it("still reads table column properties after a flushed row delete, while the table end throws", () => {
    const raw = sheetAfterFlushedDataRowDelete();

    const table = raw.table(tableId111);
    expect(table.tableId).toBe("fake-table-111");
    expect(table.startRowIndex).toBe(tableHeaderRowIndex);
    expect(table.startColumnIndex).toBe(startTableColIndex);
    expect(table.columnCount).toBeGreaterThan(0);
    const columnMeta = raw.table(tableId111).meta.column(0);
    expect(columnMeta.activeColumnType).toBe("TEXT");
    expect(columnMeta.valueValidationStrings).toEqual(["=valueConfig[Notes]"]);
    expect(columnMeta.validationConditionType).toBe("BOOLEAN");
    expect(raw.table(tableId111).isTableColIndex(0)).toBe(true);
    expect(() => table.dataRowCount).toThrow(staleRowIndexes);
    expect(() => table.growDataRowCount()).toThrow(staleRowIndexes);
  });

  it("keys declared column types by Table column when the Table starts after column A", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          table: {
            startColumnIndex: 1,
            endRowIndex: 11,
            endColumnIndex: 4,
            columnTypes: { 1: "TEXT" },
            columnValidationValues: { 1: ["=valueConfig[Notes]"] },
            columnValidationConditionTypes: { 1: "BOOLEAN" },
          },
        },
      ],
    });
    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();

    const columnMeta = raw.table(tableId111).meta.column(0);
    expect(columnMeta.activeColumnType).toBe("TEXT");
    expect(columnMeta.valueValidationStrings).toEqual(["=valueConfig[Notes]"]);
    expect(columnMeta.validationConditionType).toBe("BOOLEAN");
  });

  it("throws on per-cell value, formula, and colour writes after a flushed row delete", () => {
    const raw = sheetAfterFlushedDataRowDelete(true);
    const cell = raw.table(tableId111).row(0).cell(0);

    expect(() => cell.updateValue("painted")).toThrow(staleRowIndexes);
    expect(() => cell.updateFormula("=1")).toThrow(staleRowIndexes);
    expect(() => cell.updateBackgroundColor(lightGreen)).toThrow(
      staleRowIndexes,
    );
    expect(() => cell.addCheckboxValidation()).toThrow(staleRowIndexes);
  });

  it("throws on active-row and whole-column fills after a flushed row delete", () => {
    const raw = sheetAfterFlushedDataRowDelete(true);
    const column = raw.table(tableId111).column(0);

    expect(() => column.updateActiveCells({ value: "fill" })).toThrow(
      staleRowIndexes,
    );
    expect(() => column.updateActiveFormulas("=1")).toThrow(staleRowIndexes);
    expect(() => column.updateAllCells({ value: "fill" })).toThrow(
      staleRowIndexes,
    );
    expect(() => column.updateAllFormulas("=1")).toThrow(staleRowIndexes);
  });

  it("throws on a further data-row delete after a flushed row delete", () => {
    const raw = sheetAfterFlushedDataRowDelete();

    expect(() => raw.table(tableId111).row(2).delete()).toThrow(
      staleRowIndexes,
    );
  });

  it("refuses at the flush a sort queued after a flushed row delete", () => {
    const raw = sheetAfterFlushedDataRowDelete();

    raw
      .table(tableId111)
      .requestSortGSheet({ colIdxToSortBy: 0, sortOrder: "ASCENDING" });

    expect(() => raw.batchUpdateGSheets()).toThrow(staleRowIndexes);
  });

  it("refuses a column insert after a flushed row delete", () => {
    const raw = sheetAfterFlushedDataRowDelete();

    expect(() =>
      raw
        .sheetMeta(111)
        .insertColumnAtEnd({ columnId: "c:x:new", header: "New" }),
    ).toThrow(staleRowIndexes);
  });

  it("still reads an already-fetched cell after a flushed row delete", () => {
    const raw = sheetAfterFlushedDataRowDelete(true);

    expect(raw.table(tableId111).row(0).cell(0).valueOrEmpty()).toBe("kept");
  });

  it("leaves row indexes stale after a properties fetch that follows a flushed row delete", () => {
    const raw = sheetAfterFlushedDataRowDelete();
    raw.fetchAllSheetProperties();

    expect(raw.table(tableId111).rowIndexesAreStale).toBe(true);
    expect(() => raw.table(tableId111).dataRowCount).toThrow(staleRowIndexes);
  });

  it("clears row-index stale only when clearRowIndexStale is called, and then a write is allowed again", () => {
    const raw = sheetAfterFlushedDataRowDelete(true);
    const cell = raw.table(tableId111).row(0).cell(0);
    expect(() => cell.updateValue("painted")).toThrow(staleRowIndexes);

    raw.table(tableId111).clearRowIndexStale();

    expect(raw.table(tableId111).rowIndexesAreStale).toBe(false);
    expect(raw.table(tableId111).dataRowCount).toBe(7);
    expect(() => cell.updateValue("painted")).not.toThrow();
  });

  it("does not mark row indexes stale when a queued delete is discarded before flush", () => {
    stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).row(1).delete();
    raw.discardQueuedChanges();

    expect(raw.table(tableId111).rowIndexesAreStale).toBe(false);
    expect(() =>
      raw.table(tableId111).row(0).cell(0).updateValue("ok"),
    ).not.toThrow();
  });

  it("deletes exactly the queued rows on one sheet, so an earlier deletion can't shift a later one out from under it", () => {
    const { grid } = stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({
            4: ["r4"],
            5: ["r5"],
            6: ["r6"],
            7: ["r7"],
            8: ["r8"],
            9: ["r9"],
            10: ["r10"],
          }),
          table: { endRowIndex: 11 },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).row(1).delete();
    raw.table(tableId111).row(6).delete();
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).values({ startRowIndex: 4 })).toEqual([
      ["r4"],
      ["r6"],
      ["r7"],
      ["r8"],
      ["r9"],
    ]);
  });
});

describe("TableRaw.requestSortGSheet", () => {
  const besideStart = startTableColIndex + 3;
  const belowRowIndex = topDataRowIndex + 3;
  function stubTableWithNeighbours() {
    return stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({
            [tableHeaderRowIndex - 1]: ["a0", "a1", "", "a3", "a4"],
            [tableHeaderRowIndex]: ["ID", "Rank", "", "Code", "Qty"],
            [topDataRowIndex]: ["r1", 3, "", "c1", 9],
            [topDataRowIndex + 1]: ["r2", 1, "", "c2", 8],
            [topDataRowIndex + 2]: ["r3", 2, "", "c3", 7],
            [belowRowIndex]: ["b0", 0, "", "b3", 0],
          }),
          tables: [
            {
              tableId: "ranked",
              name: "Ranked",
              endColumnIndex: startTableColIndex + 2,
              endRowIndex: belowRowIndex,
            },
            {
              tableId: "beside",
              name: "Beside",
              startColumnIndex: besideStart,
              endColumnIndex: besideStart + 2,
              endRowIndex: belowRowIndex,
            },
          ],
        },
      ],
    });
  }
  function rankedSortQueued() {
    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw
      .table("ranked")
      .requestSortGSheet({ colIdxToSortBy: 1, sortOrder: "ASCENDING" });
    return raw;
  }

  it("sorts the Table's body only, leaving its head rows, the rows below and the Table beside it", () => {
    const { grid } = stubTableWithNeighbours();

    rankedSortQueued().batchUpdateGSheets();

    expect(
      grid.sheet(111).values({ startRowIndex: tableHeaderRowIndex - 1 }),
    ).toEqual([
      ["a0", "a1", "", "a3", "a4"],
      ["ID", "Rank", "", "Code", "Qty"],
      ["r2", 1, "", "c1", 9],
      ["r3", 2, "", "c2", 8],
      ["r1", 3, "", "c3", 7],
      ["b0", 0, "", "b3", 0],
    ]);
  });

  it("sorts only the rows a same-flush delete leaves, so the rows pulled up below stay put", () => {
    const { grid } = stubTableWithNeighbours();

    const raw = rankedSortQueued();
    raw.table("ranked").row(1).delete();
    raw.batchUpdateGSheets();

    expect(
      grid.sheet(111).values({
        startRowIndex: topDataRowIndex,
        endRowIndex: belowRowIndex,
        endColumnIndex: startTableColIndex + 2,
      }),
    ).toEqual([
      ["r3", 2],
      ["r1", 3],
      ["b0", 0],
    ]);
  });

  it("marks the sorted Table's row indexes stale, so a row write in a later flush throws", () => {
    stubTableWithNeighbours();

    const raw = rankedSortQueued();
    raw.batchUpdateGSheets();

    expect(() => raw.table("ranked").row(0).cell(1).updateValue(5)).toThrow(
      /Row indexes are stale for Table "Ranked"/,
    );
    expect(raw.table("beside").rowIndexesAreStale).toBe(false);
  });

  it("refuses at the flush a sort queued on a Table whose rows a flush has moved, naming it", () => {
    stubTableWithNeighbours();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table("ranked").row(0).delete();
    raw.batchUpdateGSheets();
    raw
      .table("ranked")
      .requestSortGSheet({ colIdxToSortBy: 1, sortOrder: "ASCENDING" });

    expect(() => raw.batchUpdateGSheets()).toThrow(
      /Row indexes are stale for Table "Ranked"/,
    );
  });
});

describe("SpreadsheetRaw.gatherRawOperation", () => {
  function rawValueWrite(rowIndex: number, colIndex: number, value: string) {
    return googleRawRequest({
      updateCells: {
        start: { sheetId: 111, rowIndex, columnIndex: colIndex },
        rows: [{ values: [{ userEnteredValue: { stringValue: value } }] }],
        fields: "userEnteredValue",
      },
    });
  }

  it("applies a raw request after every write the framework models, so it lands on the sheet those writes left", () => {
    const { grid } = stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          table: { endRowIndex: 11, endColumnIndex: 3 },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.gatherRawOperation(rawValueWrite(5, 2, "Raw"));
    raw.gatherRawOperation(rawValueWrite(8, 2, "Raw"));
    raw.table(tableId111).row(1).cell(2).updateValue("Processing...");
    raw.table(tableId111).row(5).cell(2).updateValue("Shifted up");
    raw.table(tableId111).row(4).delete();
    raw.batchUpdateGSheets();

    expect(
      grid.sheet(111).values({ startRowIndex: 5, endRowIndex: 10 }),
    ).toEqual([
      [null, null, "Raw"],
      [null, null, null],
      [null, null, null],
      [null, null, "Raw"],
      [null, null, null],
    ]);
  });

  it("discards a raw request alongside every other queued change, sending no batch update", () => {
    const { batchUpdateCount } = stubSheetsService();

    const raw = SpreadsheetRaw.init();
    raw.gatherRawOperation(
      googleRawRequest({ updateTable: { table: { tableId: "t" } } }),
    );
    raw.discardQueuedChanges();
    raw.batchUpdateGSheets();

    expect(batchUpdateCount()).toBe(0);
  });
});

describe("queued writes outlive a same-run re-fetch", () => {
  function stubTwoDataRows(topDataRow: FakeCell[] = ["r:lse:1", "live"]) {
    stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({
            0: ["c:lse:aaa", "c:lse:bbb"],
            [tableHeaderRowIndex]: ["ID", "Status"],
            [topDataRowIndex]: topDataRow,
            [topDataRowIndex + 1]: ["r:lse:2", "other"],
          }),
          table: { endRowIndex: topDataRowIndex + 2 },
        },
      ],
    });
  }

  function fetchedSpreadsheet() {
    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).topRow.gatherFetchFull();
    raw.table(tableId111).row(1).gatherFetchFull();
    raw.fetchAllGathered();
    return raw;
  }

  it("leaves a row queued for delete inactive after a re-fetch that returns it", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.table(tableId111).topRow.delete();
    raw.table(tableId111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.table(tableId111).topRow.rowIsActive()).toBe(false);
  });

  it("leaves the same row inactive when the re-fetch was a full row, so finalize backfilled", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.table(tableId111).topRow.delete();
    raw.table(tableId111).topRow.gatherFetchFull();
    expect(() => raw.fetchAllGathered()).not.toThrow();
    expect(raw.table(tableId111).topRow.rowIsActive()).toBe(false);
  });

  it("leaves the same row inactive after a full-column fetch that covers it", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.table(tableId111).topRow.delete();
    raw.table(tableId111).column(1).gatherFetchFull();
    expect(() => raw.fetchAllGathered()).not.toThrow();
    expect(raw.table(tableId111).topRow.rowIsActive()).toBe(false);
  });

  it("still supplies Table column facts from a top data row queued for delete", () => {
    stubTwoDataRows([
      "r:lse:1",
      { value: 42, isFormula: true, numberFormatType: "CURRENCY" },
    ]);

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).topRow.gatherFetchFull();
    raw.fetchAllGathered(true);
    raw.table(tableId111).topRow.delete();
    raw.table(tableId111).topRow.gatherFetchFull();
    expect(() => raw.fetchAllGathered(true)).not.toThrow();

    const column = raw.sheetMeta(111).column(1);
    expect(column.activeIsFormula).toBe(true);
    expect(column.activeNumberFormatType).toBe("CURRENCY");
    expect(column.activeTopValue).toBe(42);
  });

  it("still answers whether the top data row is blank from those facts", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.table(tableId111).topRow.delete();
    raw.table(tableId111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.table(tableId111).topDataRowIsBlank()).toBe(false);
  });

  it("keeps a queued value update after a re-fetch of that cell", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.table(tableId111).topRow.cell(1).updateValue("queued");
    raw.table(tableId111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.table(tableId111).topRow.valueOrEmpty(1)).toBe("queued");
  });

  it("keeps a queued value fill after a re-fetch of a covered cell", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.table(tableId111).column(1).updateAllCells({ value: "filled" });
    raw.table(tableId111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.table(tableId111).topRow.valueOrEmpty(1)).toBe("filled");
  });

  it("lets the most recently queued value fill win when several cover one cell", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.table(tableId111).column(1).updateAllCells({ value: "first" });
    raw.table(tableId111).column(1).updateAllCells({ value: "second" });
    raw.table(tableId111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.table(tableId111).topRow.valueOrEmpty(1)).toBe("second");
  });

  it("lets the cell's own queued value win over a covering fill", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.table(tableId111).column(1).updateAllCells({ value: "filled" });
    raw.table(tableId111).topRow.cell(1).updateValue("queued");
    raw.table(tableId111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.table(tableId111).topRow.valueOrEmpty(1)).toBe("queued");
  });

  it("takes the live value when the cell has no queued value", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.table(tableId111).topRow.cell(1).updateValue("stale local");
    raw.discardQueuedChanges();
    raw.table(tableId111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.table(tableId111).topRow.valueOrEmpty(1)).toBe("live");
  });

  it("takes the live value when only a formula update is queued", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.table(tableId111).topRow.cell(1).updateFormula("=A1");
    raw.table(tableId111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.table(tableId111).topRow.valueOrEmpty(1)).toBe("live");
  });

  it("takes the live value when only a formula fill is queued", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.table(tableId111).column(1).updateAllFormulas("=A1");
    raw.table(tableId111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.table(tableId111).topRow.valueOrEmpty(1)).toBe("live");
  });

  it("integrates live values after the flush has cleared the queue", () => {
    stubTwoDataRows();

    const raw = fetchedSpreadsheet();
    raw.table(tableId111).topRow.cell(1).updateValue("queued");
    raw.batchUpdateGSheets();
    const otherRun = fetchedSpreadsheet();
    otherRun.table(tableId111).topRow.cell(1).updateValue("live");
    otherRun.batchUpdateGSheets();
    raw.table(tableId111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.table(tableId111).topRow.valueOrEmpty(1)).toBe("live");
  });

  it("applies a value queued before the row was fetched once that row is fetched", () => {
    stubTwoDataRows();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).topRow.cell(1).updateValue("queued");
    raw.table(tableId111).topRow.gatherFetchFull();
    raw.fetchAllGathered();

    expect(raw.table(tableId111).topRow.valueOrEmpty(1)).toBe("queued");
  });

  function stubNamedTables() {
    stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({
            0: ["c:lse:aaa", "c:lse:bbb"],
            [tableHeaderRowIndex]: ["ID", "Status"],
          }),
          table: {
            name: "records",
            endRowIndex: topDataRowIndex + 1,
            columnTypes: { 1: "TEXT" },
          },
        },
        {
          sheetId: 222,
          title: "Entries",
          rows: buildGridRows({ [tableHeaderRowIndex]: ["ID"] }),
          table: { name: "entries", endRowIndex: topDataRowIndex + 1 },
        },
      ],
    });
  }

  it("keeps a queued tab title after a re-fetch of the sheet properties", () => {
    stubNamedTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).updateTitle("Renamed");
    raw.fetchAllSheetProperties();

    expect(raw.table(tableId111).title).toBe("Renamed");
  });

  it("keeps a queued Table name on the known Table and in the sheet's Tables after a re-fetch", () => {
    stubNamedTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).updateTableName("renamedRecords");
    raw.fetchAllSheetProperties();

    expect(raw.table(tableId111).name).toBe("renamedRecords");
    expect(raw.table(tableId111).tableIds()).toEqual(["fake-table-111"]);
  });

  it("keeps a queued column type after a re-fetch of the sheet properties", () => {
    stubNamedTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.sheetMeta(111).column(1).updateColumnType("DOUBLE");
    raw.fetchAllSheetProperties();

    expect(raw.sheetMeta(111).column(1).activeColumnType).toBe("DOUBLE");
  });

  it("lets the last of two queued tab titles win after a re-fetch", () => {
    stubNamedTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).updateTitle("First");
    raw.table(tableId111).updateTitle("Second");
    raw.fetchAllSheetProperties();

    expect(raw.table(tableId111).title).toBe("Second");
  });

  it("lets the last of two queued Table names win after a re-fetch", () => {
    stubNamedTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).updateTableName("firstRecords");
    raw.table(tableId111).updateTableName("secondRecords");
    raw.fetchAllSheetProperties();

    expect(raw.table(tableId111).name).toBe("secondRecords");
  });

  it("leaves another sheet's queued title alone on a re-fetch of one sheet", () => {
    stubNamedTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId222).updateTitle("Renamed");
    raw.fetchSheetUsedGrid(111);

    expect(raw.table(tableId222).title).toBe("Renamed");
  });

  it("applies a sheet's queued title to that sheet only", () => {
    stubNamedTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).updateTitle("Renamed");
    raw.fetchAllSheetProperties();

    expect(raw.table(tableId222).title).toBe("Entries");
  });

  it("integrates the live title and Table name after the flush has cleared the queue", () => {
    stubNamedTables();

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).updateTitle("Renamed");
    raw.table(tableId111).updateTableName("renamedRecords");
    raw.batchUpdateGSheets();
    const otherRun = SpreadsheetRaw.init();
    otherRun.fetchAllSheetProperties();
    otherRun.table(tableId111).updateTitle("Records");
    otherRun.table(tableId111).updateTableName("records");
    otherRun.batchUpdateGSheets();
    raw.fetchAllSheetProperties();

    expect(raw.table(tableId111).title).toBe("Records");
    expect(raw.table(tableId111).name).toBe("records");
  });
});

describe("the last queued write wins between fills and cell writes", () => {
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
          }),
          table: { endRowIndex: 6 },
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

  it("shows a fill queued after a cell write", () => {
    const { grid } = stubFilledSheet();

    const raw = fetchedColumn();
    raw.table(tableId111).row(1).cell(1).updateValue("cell");
    raw.table(tableId111).column(1).updateAllCells({ value: "filled" });
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).values(gridRanges.columnOneData)).toEqual([
      ["filled"],
      ["filled"],
    ]);
  });

  it("keeps the cell's colour under a later value-only fill", () => {
    const { grid } = stubFilledSheet();

    const raw = fetchedColumn();
    raw
      .table(tableId111)
      .row(1)
      .cell(1)
      .updateValue("cell")
      .updateBackgroundColor(lightGreen);
    raw.table(tableId111).column(1).updateAllCells({ value: "filled" });
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).rows(gridRanges.columnOneData)).toEqual([
      ["filled"],
      [{ value: "filled", backgroundColor: lightGreen }],
    ]);
  });

  it("keeps the cell's value under a later colour-only fill", () => {
    const { grid } = stubFilledSheet();

    const raw = fetchedColumn();
    raw.table(tableId111).row(1).cell(1).updateValue("cell");
    raw
      .table(tableId111)
      .column(1)
      .updateAllCells({ backgroundColor: lightGreen });
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).rows(gridRanges.columnOneData)).toEqual([
      [{ value: "old", backgroundColor: lightGreen }],
      [{ value: "cell", backgroundColor: lightGreen }],
    ]);
  });

  it("drops a queued cell formula under a later value fill", () => {
    const { grid } = stubFilledSheet();

    const raw = fetchedColumn();
    raw.table(tableId111).row(1).cell(1).updateFormula("=1+1");
    raw.table(tableId111).column(1).updateAllCells({ value: "filled" });
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).rows(gridRanges.columnOneData)).toEqual([
      ["filled"],
      ["filled"],
    ]);
  });

  it("reads before the flush what the grid shows after it, across a re-fetch", () => {
    const { grid } = stubFilledSheet();

    const raw = fetchedColumn();
    raw.table(tableId111).row(1).cell(1).updateValue("cell");
    raw.table(tableId111).column(1).updateAllCells({ value: "filled" });
    raw.table(tableId111).row(1).gatherFetchFull();
    raw.fetchAllGathered();
    const readBeforeFlush = raw.table(tableId111).row(1).valueOrEmpty(1);
    raw.batchUpdateGSheets();

    expect([readBeforeFlush, grid.sheet(111).cell(5, 1)]).toEqual([
      "filled",
      "filled",
    ]);
  });

  it("leaves a cell write on a row appended after the fill alone", () => {
    const { grid } = stubFilledSheet();

    const raw = fetchedColumn();
    raw.table(tableId111).column(1).updateAllCells({ value: "filled" });
    raw.table(tableId111).appendDataRow().cell(1).updateValue("appended");
    raw.batchUpdateGSheets();

    expect(
      grid.sheet(111).values({ ...gridRanges.columnOneData, endRowIndex: 7 }),
    ).toEqual([["filled"], ["filled"], ["appended"]]);
  });
});

describe("SpreadsheetRaw.discardQueuedChanges", () => {
  it("sends no batch update for changes queued before the discard", () => {
    const { batchUpdateCount } = stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).row(1).delete();
    raw.discardQueuedChanges();
    raw.batchUpdateGSheets();

    expect(batchUpdateCount()).toBe(0);
    expect(raw.table(tableId111).rowIndexesAreStale).toBe(false);
    expect(raw.table(tableId111).dataRowCount).toBe(7);
  });

  it("empties the spreadsheet and per-sheet write queues, so a later flush sends no batch update", () => {
    const { batchUpdateCount } = stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).row(1).delete();
    raw.table(tableId111).requestSortGSheet({
      colIdxToSortBy: 0,
      sortOrder: "ASCENDING",
    });
    raw.gatherRawOperation(
      googleRawRequest({ updateTable: { table: { tableId: "t" } } }),
    );
    raw.discardQueuedChanges();
    raw.batchUpdateGSheets();

    expect(batchUpdateCount()).toBe(0);
  });

  it("still applies changes queued after the discard, so a failure handler can report status", () => {
    const { grid } = stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({ 4: ["r4"], 5: ["r5"] }),
          table: { endRowIndex: 11 },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).row(1).delete();
    raw.discardQueuedChanges();
    raw.table(tableId111).appendDataRow();
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).values({ startRowIndex: 4 })).toEqual([
      ["r4"],
      ["r5"],
    ]);
    expect(firstTableEndRowIndex(grid, 111)).toBe(12);
  });
});

describe("CellRaw.updateValue", () => {
  it("writes to a row that was never fetched, since a write needs no fetched state", () => {
    const { grid } = stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).row(1).cell(2).updateValue("Processing...");
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).values()).toEqual([
      [null, null, null],
      [null, null, null],
      [null, null, null],
      [null, null, null],
      [null, null, null],
      [null, null, "Processing..."],
    ]);
  });

  it("leaves an unfetched row unreadable, so a forgotten fetch still fails loudly on read", () => {
    stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    const cell = raw.table(tableId111).row(1).cell(2);
    cell.updateValue("Processing...");

    expect(() => cell.valueOrEmpty()).toThrowError(/No value is set/);
  });

  it("throws for a data row past the table's last row rather than writing off the grid", () => {
    stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();

    expect(() =>
      raw.table(tableId111).row(7).cell(2).updateValue("x"),
    ).toThrowError(/past the last row/);
  });

  it("reflects the write in row state when the row was fetched, so a later read sees it", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({
            0: ["c:lse:aaa", "c:lse:bbb"],
            4: ["r:lse:1", "old"],
          }),
          table: { endRowIndex: 5 },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).row(0).gatherFetchFull();
    raw.fetchAllGathered();
    const cell = raw.table(tableId111).row(0).cell(1);
    cell.updateValue("new");

    expect(cell.valueOrEmpty()).toBe("new");
  });
});

describe("CellRaw.updateBackgroundColor", () => {
  it("colours the cell and leaves its value alone", () => {
    const { grid } = stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({ 5: [null, null, "existing"] }),
          table: { endRowIndex: 11 },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).row(1).cell(2).updateBackgroundColor(lightGreen);
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).cell(5, 2)).toEqual({
      value: "existing",
      backgroundColor: lightGreen,
    });
  });

  it("writes the value queued before the colour on the same cell", () => {
    const { grid } = stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).row(1).cell(2).updateValue("2026-09-05 10:00:00");
    raw.table(tableId111).row(1).cell(2).updateBackgroundColor(lightGreen);
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).cell(5, 2)).toEqual({
      value: "2026-09-05 10:00:00",
      backgroundColor: lightGreen,
    });
  });

  it("leaves the cell unreadable, since the read path never fetches colour", () => {
    stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    const cell = raw.table(tableId111).row(1).cell(2);
    cell.updateBackgroundColor(lightGreen);

    expect(cell.isActive).toBe(false);
    expect(() => cell.valueOrEmpty()).toThrowError(/No value is set/);
  });
});

describe("CellRaw.addCheckboxValidation", () => {
  it("makes that one cell a checkbox", () => {
    const { grid } = stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records", table: { endRowIndex: 11 } }],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).row(1).cell(2).addCheckboxValidation();
    raw.batchUpdateGSheets();

    const checkbox = { value: null, dataValidationConditionType: "BOOLEAN" };
    expect(grid.sheet(111).rows(threeByThreeAround(5, 2))).toEqual([
      [null, null, null],
      [null, checkbox, null],
      [null, null, null],
    ]);
  });
});

// A mis-wired accessor still type-checks; the instance checks catch it.
