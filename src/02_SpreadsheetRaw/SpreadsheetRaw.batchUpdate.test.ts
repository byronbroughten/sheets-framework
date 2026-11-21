import { describe, expect, it, vi } from "vitest";

import {
  installedRawSource,
  type LocalWriteOperation,
} from "../00_Source/RawSource/RawSource";
import { SheetIndex } from "../00_Source/RawSource/SheetIndex";
import {
  buildGridRows,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { SpreadsheetRaw } from "./SpreadsheetRaw";
import {
  firstTableEndRowIndex,
  lightGreen,
  startTableColIndex,
  tableHeaderRowIndex,
  tableId111,
  tableId222,
  topDataRowIndex,
} from "./spreadsheetRawTestSupport";

describe("SpreadsheetRaw.batchUpdateGSheets", () => {
  it("sends nothing for a write queued before the sheet's Table was fetched, once discarded", () => {
    const { batchUpdateCount } = stubSheetsService({
      sheets: [{ sheetId: 111, title: "Records" }],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(111).table(tableId111).queueSort({
      colIdxToSortBy: 0,
      sortOrder: "ASCENDING",
    });
    raw.discardQueuedChanges();
    raw.batchUpdateGSheets();

    expect(batchUpdateCount()).toBe(0);
  });

  it("replaces the conditional-format rule at the queued index in place on flush", () => {
    const { grid } = stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          conditionalFormats: [
            {
              ranges: [
                { sheetId: 111, startColumnIndex: 0, endColumnIndex: 1 },
              ],
            },
            {
              ranges: [
                { sheetId: 111, startColumnIndex: 1, endColumnIndex: 2 },
              ],
            },
          ],
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.sheet(111).updateConditionalFormatRule(1, {
      kind: "boolean",
      ranges: [
        {
          sheetId: 111,
          startRowIndex: SheetIndex.row(2),
          endRowIndex: SheetIndex.row(4),
          startColumnIndex: SheetIndex.col(3),
          endColumnIndex: SheetIndex.col(4),
        },
      ],
      condition: { type: "CUSTOM_FORMULA", formula: "=$D3=TRUE" },
      format: { backgroundColor: lightGreen },
    });
    raw.batchUpdateGSheets();

    const rules = grid.sheet(111).conditionalFormats;
    expect(rules.map(({ ranges }) => ranges?.[0]?.startColumnIndex)).toEqual([
      0, 3,
    ]);
    expect(rules[1]?.booleanRule?.condition?.values).toEqual([
      { userEnteredValue: "=$D3=TRUE" },
    ]);
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

    describe("formula columns", () => {
      const formula = { value: "=1+1", isFormula: true } as const;
      function stubFormulaTable({ hasRowBelow }: { hasRowBelow: boolean }) {
        return stubSheetsService({
          sheets: [
            {
              sheetId: 111,
              title: "Records",
              rows: buildGridRows({
                [tableHeaderRowIndex]: ["ID", "A", "B", "Note", "C"],
                [topDataRowIndex]: ["r1", formula, formula, "n1", formula],
                ...(hasRowBelow ? { [topDataRowIndex + 1]: [] } : {}),
              }),
              table: { endRowIndex: topDataRowIndex + 1 },
            },
          ],
        });
      }
      // A test of the request mapping itself: the fake grows a Table on the grid's edge on appendDimension alone, so only the requests show the insert was dropped.
      function sentGrowthOperations(
        appendCount: number,
        formulaColIndexes: number[],
      ): LocalWriteOperation[] {
        const flush = vi.spyOn(installedRawSource(), "flush");
        const raw = SpreadsheetRaw.init();
        raw.fetchAllSheetProperties();
        for (let count = 0; count < appendCount; count++) {
          raw.table(tableId111).appendDataRow(formulaColIndexes);
        }
        raw.batchUpdateGSheets();
        return (flush.mock.calls[0]?.[0] ?? []).filter(
          (operation) =>
            operation.kind === "insertRange" ||
            (operation.kind === "copyPaste" &&
              operation.pasteType === "PASTE_FORMULA"),
        );
      }
      function rowsRange(
        startRowIndex: number,
        endRowIndex: number,
        startColIndex: number,
        endColIndex: number,
      ) {
        return {
          sheetId: 111,
          startRowIndex,
          endRowIndex,
          startColumnIndex: startTableColIndex + startColIndex,
          endColumnIndex: startTableColIndex + endColIndex,
        };
      }
      function formulaPaste(startColIndex: number, endColIndex: number) {
        return {
          kind: "copyPaste",
          source: rowsRange(
            topDataRowIndex,
            topDataRowIndex + 1,
            startColIndex,
            endColIndex,
          ),
          destination: rowsRange(
            topDataRowIndex + 1,
            topDataRowIndex + 3,
            startColIndex,
            endColIndex,
          ),
          pasteType: "PASTE_FORMULA",
        };
      }

      it("inserts the new rows and pastes formulas over each run of the named columns when the Table has room below", () => {
        stubFormulaTable({ hasRowBelow: true });

        expect(sentGrowthOperations(2, [1, 2, 4])).toEqual([
          {
            kind: "insertRange",
            range: rowsRange(topDataRowIndex + 1, topDataRowIndex + 3, 0, 5),
            shiftDimension: "ROWS",
          },
          formulaPaste(1, 3),
          formulaPaste(4, 5),
        ]);
      });

      it("sends no insert when the Table ends on the grid's last row, and still pastes its formulas", () => {
        stubFormulaTable({ hasRowBelow: false });

        expect(sentGrowthOperations(2, [1, 2, 4])).toEqual([
          formulaPaste(1, 3),
          formulaPaste(4, 5),
        ]);
      });

      it("pastes no formula for an append that names no formula columns", () => {
        stubFormulaTable({ hasRowBelow: true });

        expect(sentGrowthOperations(1, [])).toEqual([
          expect.objectContaining({ kind: "insertRange" }),
        ]);
      });
    });

    const leftStart = startTableColIndex;
    const rightStart = startTableColIndex + 3;
    // A grid row below both Tables keeps them off the grid's edge, which Sheets grows every Table on.
    function stubSideBySideTables({ hasRowBelow = false } = {}) {
      return stubSheetsService({
        sheets: [
          {
            sheetId: 111,
            title: "Records",
            rows: buildGridRows({
              [tableHeaderRowIndex]: ["ID", "Name", "", "Code", "Qty"],
              [topDataRowIndex]: ["r1", "a", "loose", "c1", 5],
              [topDataRowIndex + 1]: ["r2", "b", "", "c2", 6],
              ...(hasRowBelow ? { [topDataRowIndex + 2]: [] } : {}),
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
      const { grid } = stubSideBySideTables({ hasRowBelow: true });

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

    it("leaves a side-by-side neighbour on the grid's last row a blank row longer, as live", () => {
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
        topDataRowIndex + 3,
      ]);
    });

    describe("a second append to a neighbour the grid grew", () => {
      const neighbourEnd = rightStart + 2;
      function appendToRight(raw: SpreadsheetRaw): void {
        raw.table("right").appendDataRow().updateValue(0, "c3");
        raw.batchUpdateGSheets();
      }

      it("lands below the blank row an append to the other Table grew it by", () => {
        const { grid } = stubSideBySideTables();

        const raw = SpreadsheetRaw.init();
        raw.fetchAllSheetProperties();
        raw.table("left").appendDataRow().updateValue(0, "r3");
        raw.batchUpdateGSheets();
        appendToRight(raw);

        expect(
          grid.sheet(111).values({ startRowIndex: topDataRowIndex }),
        ).toEqual([
          ["r1", "a", "loose", "c1", 5],
          ["r2", "b", "", "c2", 6],
          ["r3", null, null, null, null],
          [null, null, null, "c3", null],
        ]);
        // The second append ends on the grid's edge too, so it grows "left" by a blank row in turn.
        expect(tableEndRowIndexes(grid)).toEqual([
          topDataRowIndex + 4,
          topDataRowIndex + 4,
        ]);
      });

      it("lands below the blank row a new Table placed past the grid's edge grew it by", () => {
        const { grid } = stubSideBySideTables();

        const raw = SpreadsheetRaw.init();
        raw.fetchAllSheetProperties();
        raw.createTable({
          name: "Payment Log",
          placement: { kind: "besideTables", sheetGid: 111 },
          idPrefix: "pay",
          columns: [{ header: "Payee" }],
          seedRows: [["Ann"], ["Bo"], ["Cy"]],
        });
        raw.batchUpdateGSheets();
        appendToRight(raw);

        expect(
          grid.sheet(111).values({
            startRowIndex: topDataRowIndex,
            endColumnIndex: neighbourEnd,
          }),
        ).toEqual([
          ["r1", "a", "loose", "c1", 5],
          ["r2", "b", "", "c2", 6],
          [null, null, null, null, null],
          [null, null, null, "c3", null],
        ]);
        expect(tableEndRowIndexes(grid).slice(0, 2)).toEqual([
          topDataRowIndex + 4,
          topDataRowIndex + 4,
        ]);
      });

      it("lands below the blank row a new Table queued in the same flush grew it by", () => {
        const { grid } = stubSideBySideTables();

        const raw = SpreadsheetRaw.init();
        raw.fetchAllSheetProperties();
        raw.createTable({
          name: "Payment Log",
          placement: { kind: "besideTables", sheetGid: 111 },
          idPrefix: "pay",
          columns: [{ header: "Payee" }],
          seedRows: [["Ann"], ["Bo"], ["Cy"]],
        });
        appendToRight(raw);

        expect(
          grid.sheet(111).values({
            startRowIndex: topDataRowIndex,
            endColumnIndex: neighbourEnd,
          }),
        ).toEqual([
          ["r1", "a", "loose", "c1", 5],
          ["r2", "b", "", "c2", 6],
          [null, null, null, null, null],
          [null, null, null, "c3", null],
        ]);
        expect(tableEndRowIndexes(grid).slice(0, 2)).toEqual([
          topDataRowIndex + 4,
          topDataRowIndex + 4,
        ]);
        expect(grid.sheet(111).rowCount).toBe(topDataRowIndex + 4);
      });

      it("leaves a Table not on the grid's last row its range", () => {
        stubSideBySideTables({ hasRowBelow: true });

        const raw = SpreadsheetRaw.init();
        raw.fetchAllSheetProperties();
        raw.table("left").appendDataRow().updateValue(0, "r3");
        raw.batchUpdateGSheets();

        expect(raw.table("right").dataRowCount).toBe(2);
      });

      it("holds the grown row as the grid does: the last row's formula carried down, every other cell blank", () => {
        const formula = { value: "=1+1", isFormula: true } as const;
        const { grid } = stubSheetsService({
          sheets: [
            {
              sheetId: 111,
              title: "Records",
              rows: buildGridRows({
                [tableHeaderRowIndex]: ["ID", "Name", "", "Code", "Sum"],
                [topDataRowIndex]: ["r1", "a", "", "c1", formula],
                [topDataRowIndex + 1]: ["r2", "b", "", "c2", formula],
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

        const raw = SpreadsheetRaw.init();
        raw.fetchAllSheetProperties();
        raw.table("right").row(0).gatherFetchFull();
        raw.table("right").row(1).gatherFetchFull();
        raw.fetchAllGathered();
        raw.table("left").appendDataRow().updateValue(0, "r3");
        raw.batchUpdateGSheets();

        expect(
          grid.sheet(111).values({
            startRowIndex: topDataRowIndex + 2,
            startColumnIndex: rightStart,
            endColumnIndex: rightStart + 2,
          }),
        ).toEqual([[null, "=1+1"]]);
        const grown = raw.table("right").row(2);
        expect([grown.valueOrEmpty(0), grown.valueOrEmpty(1)]).toEqual([
          "",
          "=1+1",
        ]);
      });
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

      it("lands a checkbox, a conditional-format rule and a column protection queued on a lower Table at its rows after growth, and leaves a whole-sheet protection whole", () => {
        const { grid } = stubStackedTables();

        const raw = SpreadsheetRaw.init();
        raw.fetchAllSheetProperties();
        raw.table("top").appendDataRow().updateValue(0, "t3");
        const lower = raw.table("lower");
        lower.row(1).cell(1).addCheckboxValidation();
        lower.column(0).addConditionalFormatRule({
          condition: { type: "NUMBER_EQ", value: true },
          format: { backgroundColor: lightGreen },
        });
        lower.column(1).addEditWarning({ description: "lower name" });
        raw.sheet(111).addEditWarningWholeSheet({ description: "whole" });
        raw.batchUpdateGSheets();

        const lowerBody = {
          sheetId: 111,
          startRowIndex: lowerHeaderRowIndex + 2,
          endRowIndex: lowerHeaderRowIndex + 4,
        };
        expect(
          grid.sheet(111).rows({
            ...lowerBody,
            startColumnIndex: startTableColIndex + 1,
            endColumnIndex: startTableColIndex + 2,
          }),
        ).toEqual([
          ["x"],
          [{ value: "y", dataValidationConditionType: "BOOLEAN" }],
        ]);
        expect(
          grid.sheet(111).conditionalFormats.map(({ ranges }) => ranges),
        ).toEqual([
          [
            {
              ...lowerBody,
              startColumnIndex: startTableColIndex,
              endColumnIndex: startTableColIndex + 1,
            },
          ],
        ]);
        expect(
          grid.sheet(111).protectedRanges.map(({ description, range }) => ({
            description,
            range,
          })),
        ).toEqual([
          { description: "whole", range: { sheetId: 111 } },
          {
            description: "lower name",
            range: {
              ...lowerBody,
              startColumnIndex: startTableColIndex + 1,
              endColumnIndex: startTableColIndex + 2,
            },
          },
        ]);
      });

      it("lands a conditional-format rule and a column protection queued on a lower Table at its rows after a same-batch row delete above it", () => {
        const { grid } = stubStackedTables();

        const raw = SpreadsheetRaw.init();
        raw.fetchAllSheetProperties();
        raw.table("top").row(0).delete();
        const lower = raw.table("lower");
        lower.column(0).addConditionalFormatRule({
          condition: { type: "NUMBER_EQ", value: true },
          format: { backgroundColor: lightGreen },
        });
        lower.column(1).addEditWarning({ description: "lower name" });
        raw.batchUpdateGSheets();

        const lowerBody = {
          sheetId: 111,
          startRowIndex: lowerHeaderRowIndex,
          endRowIndex: lowerHeaderRowIndex + 2,
        };
        expect(
          grid.sheet(111).conditionalFormats.map(({ ranges }) => ranges),
        ).toEqual([
          [
            {
              ...lowerBody,
              startColumnIndex: startTableColIndex,
              endColumnIndex: startTableColIndex + 1,
            },
          ],
        ]);
        expect(
          grid.sheet(111).protectedRanges.map(({ range }) => range),
        ).toEqual([
          {
            ...lowerBody,
            startColumnIndex: startTableColIndex + 1,
            endColumnIndex: startTableColIndex + 2,
          },
        ]);
      });

      it("sorts a lower Table at its rows after a same-batch row delete above it, lands its fill and rule there too, and leaves a Table beside it in place", () => {
        const { grid } = stubStackedTables();

        const raw = SpreadsheetRaw.init();
        raw.fetchAllSheetProperties();
        raw.table("top").row(0).delete();
        const lower = raw.table("lower");
        lower.row(0).cell(1).updateValue("w");
        lower.queueSort({ colIdxToSortBy: 1, sortOrder: "DESCENDING" });
        lower.column(0).addConditionalFormatRule({
          condition: { type: "NUMBER_EQ", value: true },
          format: { backgroundColor: lightGreen },
        });
        raw
          .table("aside")
          .queueSort({ colIdxToSortBy: 1, sortOrder: "DESCENDING" });
        raw.batchUpdateGSheets();

        expect(tableRows(grid)).toEqual({
          top: [tableHeaderRowIndex, topDataRowIndex + 1],
          lower: [lowerHeaderRowIndex - 1, lowerHeaderRowIndex + 2],
          aside: [lowerHeaderRowIndex, lowerHeaderRowIndex + 3],
        });
        expect(
          grid.sheet(111).values({
            startRowIndex: lowerHeaderRowIndex,
            endRowIndex: lowerHeaderRowIndex + 2,
            endColumnIndex: startTableColIndex + 2,
          }),
        ).toEqual([
          ["b2", "y"],
          ["b1", "w"],
        ]);
        expect(
          grid.sheet(111).values({
            startRowIndex: lowerHeaderRowIndex + 1,
            endRowIndex: lowerHeaderRowIndex + 3,
            startColumnIndex: asideStart,
            endColumnIndex: asideStart + 2,
          }),
        ).toEqual([
          ["c2", 6],
          ["c1", 5],
        ]);
        expect(
          grid.sheet(111).conditionalFormats.map(({ ranges }) => ranges),
        ).toEqual([
          [
            {
              sheetId: 111,
              startRowIndex: lowerHeaderRowIndex,
              endRowIndex: lowerHeaderRowIndex + 2,
              startColumnIndex: startTableColIndex,
              endColumnIndex: startTableColIndex + 1,
            },
          ],
        ]);
      });

      it("deletes a lower Table's own row before the delete above pulls it up, and fills the row it keeps", () => {
        const { grid } = stubStackedTables();

        const raw = SpreadsheetRaw.init();
        raw.fetchAllSheetProperties();
        raw.table("top").row(0).delete();
        const lower = raw.table("lower");
        lower.row(0).delete();
        lower.row(1).cell(1).updateValue("z");
        lower.queueSort({ colIdxToSortBy: 1, sortOrder: "ASCENDING" });
        raw.batchUpdateGSheets();

        expect(tableRows(grid)).toMatchObject({
          lower: [lowerHeaderRowIndex - 1, lowerHeaderRowIndex + 1],
        });
        expect(
          grid.sheet(111).values({
            startRowIndex: lowerHeaderRowIndex - 1,
            endRowIndex: lowerHeaderRowIndex + 1,
            endColumnIndex: startTableColIndex + 2,
          }),
        ).toEqual([
          ["ID", "Name"],
          ["b2", "z"],
        ]);
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
    const { profile } = raw.table(tableId111).column(0);
    expect(profile.columnType).toBe("TEXT");
    expect(profile.valueValidationStrings).toEqual(["=valueConfig[Notes]"]);
    expect(profile.validationConditionType).toBe("BOOLEAN");
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

    const { profile } = raw.table(tableId111).column(0);
    expect(profile.columnType).toBe("TEXT");
    expect(profile.valueValidationStrings).toEqual(["=valueConfig[Notes]"]);
    expect(profile.validationConditionType).toBe("BOOLEAN");
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

  it("throws on working-row and whole-column fills after a flushed row delete", () => {
    const raw = sheetAfterFlushedDataRowDelete(true);
    const column = raw.table(tableId111).column(0);

    expect(() => column.updateWorkingCells({ value: "fill" })).toThrow(
      staleRowIndexes,
    );
    expect(() => column.updateWorkingFormulas("=1")).toThrow(staleRowIndexes);
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
      .queueSort({ colIdxToSortBy: 0, sortOrder: "ASCENDING" });

    expect(() => raw.batchUpdateGSheets()).toThrow(staleRowIndexes);
  });

  it("refuses a column insert after a flushed row delete", () => {
    const raw = sheetAfterFlushedDataRowDelete();

    expect(() =>
      raw
        .table(tableId111)
        .appendColumn({ columnId: "c:x:new", header: "New" }),
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
