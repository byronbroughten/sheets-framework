import { Val } from "@byronbroughten/utils/val";
import { describe, expect, it } from "vitest";

import { googleRawRequest } from "../00_Source/GoogleSheets/GoogleSheetsAPI";
import { headRows } from "../01_SpreadsheetSchema/headRows";
import {
  buildGridRows,
  type FakeCell,
  type FakeSheetProperties,
  type FakeTable,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { type ColumnProfileRaw } from "./ColumnProfileRaw";
import { SpreadsheetRaw } from "./SpreadsheetRaw";
import {
  colIdRowIndex,
  expectedOrigin,
  fetchedRaw,
  gridRanges,
  itemGid,
  itemTableId,
  lightGreen,
  ownColumnId,
  startTableColIndex,
  tableEndRowIndex,
  tableHeaderRowIndex,
  tableId111,
  thrownMessage,
  topDataRowIndex,
} from "./spreadsheetRawTestSupport";

describe("ColumnProfileRaw sampled facts", () => {
  const tableEndRow = topDataRowIndex + 1;

  function stubSheetWithTopDataRow(
    topDataRow: FakeCell[],
    absence?: "rowsWithNoGridData" | "rowsWithNoGridBlock",
  ) {
    stubSheetsService({
      sheets: [
        {
          sheetId: itemGid,
          title: "Item",
          rows: buildGridRows({
            0: ["c:itm:aaa", "c:itm:bbb"],
            [tableHeaderRowIndex]: ["Purchase Price", "Notes"],
            [topDataRowIndex]: topDataRow,
          }),
          ...(absence ? { [absence]: [topDataRowIndex] } : {}),
          table: { endRowIndex: tableEndRow },
        },
      ],
    });
  }

  function fetchedItemColumnProfile(colIndex: number): ColumnProfileRaw {
    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(itemTableId).topRow.gatherFetchFull();
    raw.fetchAllGathered(true);
    return raw.table(itemTableId).column(colIndex).profile;
  }

  function expectBlankFacts(column: ColumnProfileRaw): void {
    expect(column.isFormula).toBe(false);
    expect(column.numberFormatType).toBeUndefined();
    expect(column.topValue).toBe("");
    expect(column.topFormula).toBeUndefined();
  }

  it("reports blank facts for a top data row returned without any cell data", () => {
    stubSheetWithTopDataRow([], "rowsWithNoGridData");

    expectBlankFacts(fetchedItemColumnProfile(0));
  });

  it("reports blank facts for a top data row returned as no grid block at all", () => {
    stubSheetWithTopDataRow([], "rowsWithNoGridBlock");

    expectBlankFacts(fetchedItemColumnProfile(0));
  });

  it("reports the same facts an empty cell inside a returned row produces", () => {
    stubSheetWithTopDataRow([null, "a note"]);

    expectBlankFacts(fetchedItemColumnProfile(0));
  });

  it("keeps the facts the payload supplied rather than seeding over them", () => {
    stubSheetWithTopDataRow([
      { value: 42, isFormula: true, numberFormatType: "CURRENCY" },
    ]);

    const column = fetchedItemColumnProfile(0);

    expect(column.isFormula).toBe(true);
    expect(column.numberFormatType).toBe("CURRENCY");
    expect(column.topValue).toBe(42);
  });

  it("gives a dry run the top data row's formula text", () => {
    stubSheetWithTopDataRow([
      { value: "=SINGLE(item[Purchase Price])", isFormula: true },
      "a note",
    ]);

    expect(fetchedItemColumnProfile(0).topFormula).toBe(
      "=SINGLE(item[Purchase Price])",
    );
    expect(fetchedItemColumnProfile(1).topFormula).toBeUndefined();
  });

  it("reports blank facts for a full-column fetch of a wholly blank column", () => {
    stubSheetWithTopDataRow([], "rowsWithNoGridData");

    const raw = SpreadsheetRaw.init();
    raw.table(itemTableId).gatherFetchProperties();
    raw.fetchAllGathered();
    raw.table(itemTableId).column(1).gatherFetchFull();
    raw.fetchAllGathered(true);

    expectBlankFacts(raw.table(itemTableId).column(1).profile);
  });

  it("reads a specifically fetched cell omitted from the payload as empty, not unfetched", () => {
    stubSheetWithTopDataRow([], "rowsWithNoGridData");

    const raw = SpreadsheetRaw.init();
    raw.table(itemTableId).gatherFetchProperties();
    raw.fetchAllGathered();
    const cell = raw.table(itemTableId).row(0).cell(0);
    cell.gatherFetchRange();
    raw.fetchAllGathered();

    expect(cell.valueOrEmpty()).toBe("");
  });

  it("throws naming the sheet and the missing fetch for a column nothing fetched", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: itemGid,
          title: "Item",
          rows: buildGridRows({
            [colIdRowIndex]: [ownColumnId(itemGid)],
            [tableHeaderRowIndex]: ["Purchase Price"],
          }),
          table: { endRowIndex: tableEndRow },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.table(itemTableId).gatherFetchProperties();
    raw.fetchAllGathered();

    const message = thrownMessage(
      () => raw.table(itemTableId).column(0).profile.isFormula,
    );
    expect(message).toContain(`"Item" (gid ${itemGid})`);
    expect(message).toMatch(/top data row/);
  });

  it("writes no facts for a grid column outside the table", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: itemGid,
          title: "Item",
          rows: buildGridRows({
            0: ["c:itm:aaa"],
            [tableHeaderRowIndex]: ["Purchase Price"],
            [topDataRowIndex]: [100000, "outside the table"],
          }),
          table: { endRowIndex: tableEndRow, endColumnIndex: 1 },
        },
      ],
    });

    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(itemTableId).topRow.gatherFetchFull();
    raw.fetchAllGathered(true);

    expect(raw.table(itemTableId).column(0).profile.topValue).toBe(100000);
    expect(
      () => raw.table(itemTableId).column(1).profile.topValue,
    ).toThrowError(/No sampled facts/);
  });
});

describe("TableRaw.appendColumn", () => {
  const newColumn = { startColumnIndex: 3, endColumnIndex: 4 } as const;

  function stubThreeColumnTable(
    tableOverrides: Partial<FakeTable> = {},
    rightmostCells: Record<number, FakeCell> = {},
  ) {
    const rows = buildGridRows({
      [colIdRowIndex]: ["c:lse:aaa", "c:lse:bbb", "c:lse:ccc"],
      [tableHeaderRowIndex]: ["ID", "Left", "Right"],
      [topDataRowIndex]: ["r:lse:1", "left", "right"],
    });
    Object.entries(rightmostCells).forEach(([rowIndex, cell]) => {
      const row = Val.assert(rows[Number(rowIndex)], "fixture row");
      row[2] = cell;
    });
    return stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows,
          table: { endRowIndex: 11, endColumnIndex: 3, ...tableOverrides },
        },
      ],
    });
  }

  function tableColumnNames(
    grid: ReturnType<typeof stubSheetsService>["grid"],
  ): (string | undefined)[] | undefined {
    return grid
      .sheet(111)
      .tables[0]?.columnProperties?.map((column) => column.columnName);
  }

  it("grows the Table by one column holding the new header, column ID and group heading", () => {
    const { grid } = stubThreeColumnTable();

    const raw = fetchedRaw();
    const insertedIndex = raw.table(tableId111).appendColumn({
      columnId: "c:lse:ddd",
      header: "New",
      groupHeading1: "Group",
    });
    raw.batchUpdateGSheets();

    expect(insertedIndex).toBe(3);
    expect(raw.table(tableId111).columnCount).toBe(4);
    expect(grid.sheet(111).tables[0]?.range?.endColumnIndex).toBe(4);
    expect(tableColumnNames(grid)).toEqual(["ID", "Left", "Right", "New"]);
    expect(
      grid
        .sheet(111)
        .values({ ...newColumn, endRowIndex: topDataRowIndex + 1 }),
    ).toEqual([["c:lse:ddd"], ["Group"], [null], ["New"], [null]]);
  });

  it("leaves row indexes fresh and every column writable after the insert", () => {
    stubThreeColumnTable();

    const raw = fetchedRaw();
    const insertedIndex = raw
      .table(tableId111)
      .appendColumn({ columnId: "c:lse:ddd", header: "New" });
    raw.batchUpdateGSheets();

    expect(raw.table(tableId111).rowIndexesAreStale).toBe(false);
    expect(raw.table(tableId111).dataRowCount).toBe(7);
    expect(() =>
      raw.table(tableId111).row(1).cell(2).updateValue("kept"),
    ).not.toThrow();
    expect(() =>
      raw.table(tableId111).row(1).cell(insertedIndex).updateValue("new"),
    ).not.toThrow();
  });

  it("lands two inserts on one sheet side by side inside the Table, in queue order", () => {
    const { grid } = stubThreeColumnTable();

    const raw = fetchedRaw();
    const first = raw
      .table(tableId111)
      .appendColumn({ columnId: "c:lse:ddd", header: "First" });
    const second = raw
      .table(tableId111)
      .appendColumn({ columnId: "c:lse:eee", header: "Second" });
    raw.batchUpdateGSheets();

    expect([first, second]).toEqual([3, 4]);
    expect(raw.table(tableId111).columnCount).toBe(5);
    expect(grid.sheet(111).tables[0]?.range?.endColumnIndex).toBe(5);
    expect(tableColumnNames(grid)).toEqual([
      "ID",
      "Left",
      "Right",
      "First",
      "Second",
    ]);
    expect(grid.sheet(111).values(gridRanges.headerAndTopDataRow)).toEqual([
      ["ID", "Left", "Right", "First", "Second"],
      ["r:lse:1", "left", "right", null, null],
    ]);
  });

  it("starts plain beside a checkbox column, keeping none of its cells' validation, format or ticks", () => {
    const { grid } = stubThreeColumnTable(
      { columnTypes: { 2: "BOOLEAN" } },
      {
        [expectedOrigin.sheetRowIndex(headRows.index("groupHeading1"))]: {
          value: "Checks",
          backgroundColor: lightGreen,
        },
        [expectedOrigin.sheetRowIndex(headRows.index("action"))]: {
          value: true,
          dataValidationConditionType: "BOOLEAN",
        },
        [topDataRowIndex]: {
          value: true,
          dataValidationConditionType: "BOOLEAN",
          numberFormatType: "NUMBER",
        },
      },
    );

    const raw = fetchedRaw();
    raw
      .table(tableId111)
      .appendColumn({ columnId: "c:lse:ddd", header: "New" });
    raw.batchUpdateGSheets();

    expect(
      grid.sheet(111).rows({ ...newColumn, endRowIndex: topDataRowIndex + 1 }),
    ).toEqual([["c:lse:ddd"], [null], [null], ["New"], [null]]);
    expect(grid.sheet(111).tables[0]?.columnProperties?.[3]).toEqual({
      columnIndex: 3,
      columnName: "New",
    });
  });

  it("starts plain beside a dropdown column, taking neither its type nor its options", () => {
    const { grid } = stubThreeColumnTable({
      columnTypes: { 2: "DROPDOWN" },
      columnValidationConditionTypes: { 2: "ONE_OF_LIST" },
      columnValidationValues: { 2: ["left", "right"] },
    });

    const raw = fetchedRaw();
    raw
      .table(tableId111)
      .appendColumn({ columnId: "c:lse:ddd", header: "New" });
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).tables[0]?.columnProperties?.[3]).toEqual({
      columnIndex: 3,
      columnName: "New",
    });
  });

  describe("beside other Tables", () => {
    const headerRow = tableHeaderRowIndex + 2;
    const columnIdRow = headerRow - 3;
    const lastRow = headerRow + 2;
    const underHeaderRow = lastRow + 5;
    const rightStart = 3;
    function stubTablesAround() {
      return stubSheetsService({
        sheets: [
          {
            sheetId: 111,
            title: "Records",
            rows: buildGridRows({
              [columnIdRow - 1]: ["", "", "above"],
              [headerRow]: ["ID", "Name", "", "ID", "Code"],
              [headerRow + 1]: ["l1", "a", "", "r1", "x"],
              [headerRow + 2]: ["l2", "b", "", "r2", "y"],
              [lastRow + 1]: ["", "", "below"],
              [underHeaderRow]: ["ID", "Name", "Qty", "Code"],
              [underHeaderRow + 1]: ["u1", "p", 1, "q"],
            }),
            tables: [
              {
                tableId: "left",
                name: "Left",
                startRowIndex: headerRow,
                endColumnIndex: 2,
                endRowIndex: lastRow + 1,
                headRows: { 3: ["c:lft:a", "c:lft:b"] },
              },
              {
                tableId: "right",
                name: "Right",
                startRowIndex: headerRow,
                startColumnIndex: rightStart,
                endColumnIndex: rightStart + 2,
                endRowIndex: lastRow + 1,
                headRows: { 3: ["c:rgt:a", "c:rgt:b"] },
              },
              {
                tableId: "under",
                name: "Under",
                startRowIndex: underHeaderRow,
                endColumnIndex: 4,
                endRowIndex: underHeaderRow + 2,
                headRows: { 3: ["c:und:a"] },
              },
            ],
          },
        ],
      });
    }
    function tableColumns(
      grid: ReturnType<typeof stubSheetsService>["grid"],
    ): Record<string, [number | undefined, number | undefined]> {
      return Object.fromEntries(
        grid
          .sheet(111)
          .tables.map(({ tableId, range }) => [
            tableId,
            [range?.startColumnIndex, range?.endColumnIndex],
          ]),
      );
    }
    function fetchedTablesAround() {
      const { grid } = stubTablesAround();
      const raw = SpreadsheetRaw.init();
      raw.fetchAllSheetProperties();
      return { raw, grid };
    }

    it("widens the Table and moves its head rows, leaving the cells above and below it and the Table under it untouched", () => {
      const { raw, grid } = fetchedTablesAround();

      raw.table("left").appendColumn({ columnId: "c:lft:new", header: "New" });
      raw.batchUpdateGSheets();

      expect(tableColumns(grid)).toMatchObject({ left: [0, 3], under: [0, 4] });
      expect(
        grid.sheet(111).values({
          startRowIndex: columnIdRow - 1,
          endRowIndex: lastRow + 2,
          startColumnIndex: 2,
          endColumnIndex: 3,
        }),
      ).toEqual([
        ["above"],
        ["c:lft:new"],
        [null],
        [null],
        ["New"],
        [null],
        [null],
        ["below"],
      ]);
      expect(
        grid.sheet(111).values({
          startRowIndex: underHeaderRow - 3,
          endRowIndex: underHeaderRow + 2,
          endColumnIndex: 4,
        }),
      ).toEqual([
        ["c:und:a", null, null, null],
        [null, null, null, null],
        [null, null, null, null],
        ["ID", "Name", "Qty", "Code"],
        ["u1", "p", 1, "q"],
      ]);
    });

    it("shifts a Table to the right whose rows it spans, head rows included, and lands a write to it in the same batch at its new columns", () => {
      const { raw, grid } = fetchedTablesAround();

      raw.table("left").appendColumn({ columnId: "c:lft:new", header: "New" });
      raw.table("right").row(0).cell(1).updateValue("moved");
      raw.batchUpdateGSheets();

      expect(tableColumns(grid)).toMatchObject({
        right: [rightStart + 1, rightStart + 3],
      });
      expect(
        grid.sheet(111).values({
          startRowIndex: columnIdRow,
          endRowIndex: lastRow + 1,
          startColumnIndex: rightStart + 1,
          endColumnIndex: rightStart + 3,
        }),
      ).toEqual([
        ["c:rgt:a", "c:rgt:b"],
        [null, null],
        [null, null],
        ["ID", "Code"],
        ["r1", "moved"],
        ["r2", "y"],
      ]);
    });

    it("queues one protection when two side-by-side Tables protect the head row they share", () => {
      const { raw, grid } = fetchedTablesAround();

      raw.table("left").headRow("header").addEditWarning({ description: "h" });
      raw.table("right").headRow("header").addEditWarning({ description: "h" });
      raw.batchUpdateGSheets();

      expect(grid.sheet(111).protectedRanges.map(({ range }) => range)).toEqual(
        [
          {
            sheetId: 111,
            startRowIndex: headerRow,
            endRowIndex: headerRow + 1,
          },
        ],
      );
    });

    it("lands a conditional-format rule and a whole-column protection queued on a Table it shifts right at its new columns", () => {
      const { raw, grid } = fetchedTablesAround();

      raw.table("left").appendColumn({ columnId: "c:lft:new", header: "New" });
      const code = raw.table("right").column(1);
      code.addConditionalFormatRule({
        condition: { type: "NUMBER_EQ", value: true },
        format: { backgroundColor: { red: 1, green: 0, blue: 0 } },
      });
      code.addEditWarningWholeColumn({ description: "code" });
      raw.batchUpdateGSheets();

      const codeColumn = {
        sheetId: 111,
        startColumnIndex: rightStart + 2,
        endColumnIndex: rightStart + 3,
      };
      expect(
        grid.sheet(111).conditionalFormats.map(({ ranges }) => ranges),
      ).toEqual([
        [
          {
            ...codeColumn,
            startRowIndex: headerRow + 1,
            endRowIndex: lastRow + 1,
          },
        ],
      ]);
      expect(grid.sheet(111).protectedRanges.map(({ range }) => range)).toEqual(
        [{ ...codeColumn, startRowIndex: 0 }],
      );
    });

    it("lands inserts on two side-by-side Tables in one batch, each at its own end", () => {
      const { raw, grid } = fetchedTablesAround();

      raw.table("left").appendColumn({ columnId: "c:lft:new", header: "New" });
      raw
        .table("right")
        .appendColumn({ columnId: "c:rgt:new", header: "Extra" });
      raw.batchUpdateGSheets();

      expect(tableColumns(grid)).toMatchObject({
        left: [0, 3],
        right: [rightStart + 1, rightStart + 4],
      });
      expect(
        grid.sheet(111).values({
          startRowIndex: headerRow,
          endRowIndex: headerRow + 1,
          endColumnIndex: rightStart + 4,
        }),
      ).toEqual([["ID", "Name", "New", "", "ID", "Code", "Extra"]]);
    });

    describe("a neighbour the band would split", () => {
      const left: FakeTable = {
        tableId: "left",
        name: "Left",
        startRowIndex: headerRow,
        endColumnIndex: 2,
        endRowIndex: lastRow + 1,
        headRows: { 3: ["c:lft:a", "c:lft:b"] },
      };
      function neighbour(
        rows: Pick<FakeTable, "startRowIndex" | "endRowIndex">,
      ): FakeTable {
        return {
          tableId: "right",
          name: "Right",
          startColumnIndex: rightStart,
          endColumnIndex: rightStart + 2,
          headRows: { 3: ["c:rgt:a", "c:rgt:b"] },
          ...rows,
        };
      }
      function stubLeftBeside(...tables: FakeTable[]) {
        return stubSheetsService({
          sheets: [
            {
              sheetId: 111,
              title: "Records",
              rows: buildGridRows({ [headerRow]: ["ID", "Name"] }),
              tables: [left, ...tables],
            },
          ],
        });
      }
      function fetchedRawInsertingOnLeft(): SpreadsheetRaw {
        const raw = SpreadsheetRaw.init();
        raw.fetchAllSheetProperties();
        raw
          .table("left")
          .appendColumn({ columnId: "c:lft:new", header: "New" });
        return raw;
      }
      const refusal = `Inserting a column at the end of Table "Left" on sheet "Records" would shift only part of Table "Right" with its head rows. Move "Right" so that it and its head rows sit entirely within rows ${columnIdRow + 1}–${lastRow + 1}, or entirely outside them.`;

      it.each([
        [
          "whose head rows dip into the band's bottom edge",
          { startRowIndex: lastRow + 1, endRowIndex: lastRow + 3 },
        ],
        [
          "whose head rows overhang the band's top edge",
          { startRowIndex: headerRow - 1, endRowIndex: lastRow + 1 },
        ],
        [
          "that is taller than the band",
          { startRowIndex: headerRow, endRowIndex: lastRow + 3 },
        ],
      ])(
        "refuses for a neighbour %s, naming both Tables and sending nothing",
        (_, rows) => {
          const { batchUpdateCount } = stubLeftBeside(neighbour(rows));
          const raw = fetchedRawInsertingOnLeft();

          expect(() => raw.batchUpdateGSheets()).toThrow(refusal);
          expect(batchUpdateCount()).toBe(0);
        },
      );

      it("leaves a neighbour wholly below the band, head rows included, where it is", () => {
        const { grid } = stubLeftBeside(
          neighbour({ startRowIndex: lastRow + 4, endRowIndex: lastRow + 6 }),
        );
        const raw = fetchedRawInsertingOnLeft();
        raw.batchUpdateGSheets();

        expect(tableColumns(grid)).toMatchObject({
          left: [0, 3],
          right: [rightStart, rightStart + 2],
        });
      });

      it("rewords Google's refusal for a Table added after the fetch, which the app doesn't know about, naming both Tables", () => {
        const { grid } = stubLeftBeside();
        const raw = fetchedRawInsertingOnLeft();
        const operator = SpreadsheetRaw.init();
        operator.gatherRawOperation(
          googleRawRequest({
            addTable: {
              table: {
                name: "Right",
                range: {
                  sheetId: 111,
                  startRowIndex: headerRow,
                  endRowIndex: lastRow + 3,
                  startColumnIndex: rightStart,
                  endColumnIndex: rightStart + 2,
                },
              },
            },
          }),
        );
        operator.batchUpdateGSheets();

        expect(() => raw.batchUpdateGSheets()).toThrow(
          `Inserting a column at the end of Table "Left" on sheet "Records" would shift only part of Table "Right". Move "Right" so that it sits entirely within rows ${columnIdRow + 1}–${lastRow + 1}, or entirely outside them.`,
        );
        expect(tableColumns(grid)).toMatchObject({ left: [0, 2] });
      });

      it("names the Table split after growth in the same batch pushes both Tables down", () => {
        stubSheetsService({
          sheets: [
            {
              sheetId: 111,
              title: "Records",
              rows: buildGridRows({
                3: ["ID", "Name", "", "", ""],
                4: ["t1", "a"],
                9: ["ID", "Name"],
                10: ["l1", "a"],
              }),
              tables: [
                {
                  tableId: "top",
                  name: "Top",
                  startRowIndex: 3,
                  endColumnIndex: 5,
                  endRowIndex: 5,
                },
                {
                  tableId: "left",
                  name: "Left",
                  startRowIndex: 9,
                  endColumnIndex: 2,
                  endRowIndex: 11,
                  headRows: { 3: ["c:lft:a", "c:lft:b"] },
                },
              ],
            },
          ],
        });
        const raw = SpreadsheetRaw.init();
        raw.fetchAllSheetProperties();
        const operator = SpreadsheetRaw.init();
        operator.gatherRawOperation(
          googleRawRequest({
            addTable: {
              table: {
                name: "Right",
                range: {
                  sheetId: 111,
                  startRowIndex: 7,
                  endRowIndex: 12,
                  startColumnIndex: 3,
                  endColumnIndex: 5,
                },
              },
            },
          }),
        );
        operator.batchUpdateGSheets();
        raw.table("top").appendDataRow().updateValue(0, "t2");
        raw
          .table("left")
          .appendColumn({ columnId: "c:lft:new", header: "New" });

        expect(() => raw.batchUpdateGSheets()).toThrow(
          'Inserting a column at the end of Table "Left" on sheet "Records" would shift only part of Table "Right". Move "Right" so that it sits entirely within rows 8–12, or entirely outside them.',
        );
      });
    });
  });
});

describe("TableProfileRaw.columnIds", () => {
  function fetchedColumnIdTable(columnIdRow: FakeCell[]) {
    stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({
            [colIdRowIndex]: columnIdRow,
            [topDataRowIndex]: [],
          }),
          table: { endRowIndex: topDataRowIndex + 1, endColumnIndex: 2 },
        },
      ],
    });
    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table(tableId111).headRow("columnId").gatherFetchFull();
    raw.fetchAllGathered();
    return raw.table(tableId111);
  }

  it("throws when a Table column-ID cell is a number, boolean, or date", () => {
    expect(
      () => fetchedColumnIdTable(["c:lse:aaa", 42]).profile.columnIds,
    ).toThrow(/Records.*column index 1/);
    expect(
      () => fetchedColumnIdTable(["c:lse:aaa", true]).profile.columnIds,
    ).toThrow(/Records.*column index 1/);
    expect(
      () => fetchedColumnIdTable(["c:lse:aaa", 44927]).profile.columnIds,
    ).toThrow(/Records.*column index 1/);
  });

  it("treats a blank Table column-ID cell as missing rather than a type error", () => {
    expect(fetchedColumnIdTable(["c:lse:aaa", ""]).profile.columnIds).toEqual([
      "c:lse:aaa",
    ]);
  });

  it("ignores a non-string past the Table", () => {
    expect(
      fetchedColumnIdTable(["c:lse:aaa", "c:lse:bbb", 42]).profile.columnIds,
    ).toEqual(["c:lse:aaa", "c:lse:bbb"]);
  });

  it("throws from lookup by ID when a sibling Table cell is not text", () => {
    const table = fetchedColumnIdTable(["c:lse:aaa", false]);
    expect(() => table.profile.columnById("c:lse:aaa")).toThrow(
      /Records.*column index 1/,
    );
  });

  it("throws from fill-missing when a Table column-ID cell is not text", () => {
    expect(
      () =>
        fetchedColumnIdTable(["", 42]).columnResolver.colIndexesWithoutColumnId,
    ).toThrow(/Records.*column index 1/);
  });
});

describe("ColumnRaw.updateColumnType", () => {
  function stubTypedTable(
    table: Partial<NonNullable<FakeSheetProperties["table"]>> = {},
  ) {
    return stubSheetsService({
      sheets: [
        {
          sheetId: 111,
          title: "Records",
          rows: buildGridRows({
            [tableHeaderRowIndex]: ["Name", "ID", "Amount"],
          }),
          table: {
            endRowIndex: tableEndRowIndex,
            endColumnIndex: startTableColIndex + 3,
            columnTypes: { [startTableColIndex]: "TEXT" },
            ...table,
          },
        },
      ],
    });
  }

  it("sets every queued column type on the Table in the flush that appends to it, keeping each column's name", () => {
    const { grid } = stubTypedTable();
    const raw = fetchedRaw();
    raw.table(tableId111).appendDataRow();
    raw.table(tableId111).column(2).updateColumnType("DOUBLE");
    raw.table(tableId111).column(1).updateColumnType("TEXT");
    raw.batchUpdateGSheets();

    const [table] = grid.sheet(111).tables;
    expect(table?.columnProperties).toEqual([
      { columnName: "Name", columnType: "TEXT" },
      { columnIndex: 1, columnName: "ID", columnType: "TEXT" },
      { columnIndex: 2, columnName: "Amount", columnType: "DOUBLE" },
    ]);
    expect(table?.range?.endRowIndex).toBe(tableEndRowIndex + 1);
  });

  it("lets a header written in the same flush rename its column rather than be reverted by the type update", () => {
    const { grid } = stubTypedTable();
    const raw = fetchedRaw();
    raw.table(tableId111).column(2).updateColumnType("DOUBLE");
    raw
      .table(tableId111)
      .column(1)
      .headCell("header")
      .updateValue("Identifier");
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).tables[0]?.columnProperties?.[1]).toEqual({
      columnIndex: 1,
      columnName: "Identifier",
    });
  });

  it("shows a queued column type on a profile reached before the update, since the profile reads the Table's state", () => {
    stubTypedTable();
    const raw = fetchedRaw();
    const column = raw.table(tableId111).column(2);
    const { profile } = column;
    expect(profile.columnType).toBeUndefined();

    column.updateColumnType("DOUBLE");

    expect(profile.columnType).toBe("DOUBLE");
    expect(raw.table(tableId111).column(2).profile.columnType).toBe("DOUBLE");
  });

  it("keeps an untouched column's name and type through the full-list replace", () => {
    stubTypedTable();
    const raw = fetchedRaw();
    raw.table(tableId111).column(2).updateColumnType("DOUBLE");
    raw.batchUpdateGSheets();
    raw.fetchAllSheetProperties();

    expect(raw.table(tableId111).column(0).profile.columnType).toBe("TEXT");
    expect(raw.table(tableId111).column(1).profile.columnType).toBeUndefined();
    expect(raw.table(tableId111).column(2).profile.columnType).toBe("DOUBLE");
  });

  it("keeps a sibling column's type unchanged when it is one the framework does not name", () => {
    const { grid } = stubTypedTable({
      columnTypes: { [startTableColIndex]: "FUTURE_CHIP" },
    });
    const raw = fetchedRaw();
    raw.table(tableId111).column(2).updateColumnType("DOUBLE");
    raw.batchUpdateGSheets();

    expect(grid.sheet(111).tables[0]?.columnProperties).toEqual([
      { columnName: "Name", columnType: "FUTURE_CHIP" },
      { columnIndex: 1, columnName: "ID" },
      { columnIndex: 2, columnName: "Amount", columnType: "DOUBLE" },
    ]);
  });

  it("fake: a columnProperties update replaces the whole list, so a column left out loses its type", () => {
    stubTypedTable();
    const raw = fetchedRaw();
    raw.gatherRawOperation(
      googleRawRequest({
        updateTable: {
          table: {
            tableId: "fake-table-111",
            columnProperties: [
              { columnIndex: 2, columnName: "Amount", columnType: "DOUBLE" },
            ],
          },
          fields: "columnProperties",
        },
      }),
    );
    raw.batchUpdateGSheets();
    raw.fetchAllSheetProperties();

    expect(raw.table(tableId111).column(0).profile.columnType).toBeUndefined();
    expect(raw.table(tableId111).column(2).profile.columnType).toBe("DOUBLE");
  });

  it("fake: rejects a columnProperties update carrying a column with no columnName", () => {
    stubTypedTable();
    const raw = fetchedRaw();
    raw.gatherRawOperation(
      googleRawRequest({
        updateTable: {
          table: {
            tableId: "fake-table-111",
            columnProperties: [{ columnIndex: 2, columnType: "DOUBLE" }],
          },
          fields: "columnProperties",
        },
      }),
    );

    expect(() => raw.batchUpdateGSheets()).toThrow(/columnName/);
  });

  it("refuses before sending any batch update when a column on the Table has a validation rule, naming the Table and those columns", () => {
    const { batchUpdateCount } = stubTypedTable({
      columnValidationValues: { [startTableColIndex + 1]: ["a", "b"] },
      columnValidationConditionTypes: {
        [startTableColIndex + 1]: "ONE_OF_LIST",
      },
    });
    const raw = fetchedRaw();
    raw.table(tableId111).column(2).updateColumnType("DOUBLE");

    expect(() => raw.batchUpdateGSheets()).toThrow(
      /fake-table-111.*Records.*ID/,
    );
    expect(batchUpdateCount()).toBe(0);
  });

  it("refuses when the same flush inserts a column on that sheet, sending no batch update", () => {
    const { batchUpdateCount } = stubTypedTable();
    const raw = fetchedRaw();
    raw.table(tableId111).column(2).updateColumnType("DOUBLE");
    raw
      .table(tableId111)
      .appendColumn({ columnId: "c:lse:new", header: "New" });

    expect(() => raw.batchUpdateGSheets()).toThrow(/inserts a column/);
    expect(batchUpdateCount()).toBe(0);
  });

  it("refuses a second update after a flush until the Table is refetched, sending no second batch update", () => {
    const { batchUpdateCount } = stubTypedTable();
    const raw = fetchedRaw();
    raw.table(tableId111).column(2).updateColumnType("DOUBLE");
    raw.batchUpdateGSheets();
    raw.table(tableId111).column(1).updateColumnType("TEXT");

    expect(() => raw.batchUpdateGSheets()).toThrow(
      /no fetched column properties/,
    );
    expect(batchUpdateCount()).toBe(1);
  });
});
