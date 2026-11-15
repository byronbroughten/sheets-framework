import { describe, expect, it } from "vitest";

import {
  buildGridRows,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { SpreadsheetRaw } from "./SpreadsheetRaw";
import {
  colIdRowIndex,
  tableHeaderRowIndex,
  topDataRowIndex,
} from "./spreadsheetRawTestSupport";

const newGid = 777;
const existingGid = 111;
const mintedTableId = expect.stringMatching(/^tbl-[0-9a-f]{10}$/);
const noTakenIdPrefixes = { prefixesInUse: new Set<string>() };

function mintedColumnId(idPrefix: string) {
  return expect.stringMatching(new RegExp(`^c:${idPrefix}:[0-9a-zA-Z_-]{7}$`));
}

describe("SpreadsheetRaw.createTable on a tab the same flush adds", () => {
  function createdOnNewTab() {
    const fake = stubSheetsService();
    const raw = SpreadsheetRaw.init();
    const created = raw.createTable({
      name: "Payment Log",
      placement: { kind: "newTab", sheetGid: newGid, title: "Payments" },
      idPrefix: noTakenIdPrefixes,
      columns: [{ header: "Payee" }, { header: "Amount" }],
    });
    raw.batchUpdateGSheets();
    return { ...fake, created };
  }

  it("adds the tab and a Table at the expected origin holding only its listed columns and one empty body row", () => {
    const { grid } = createdOnNewTab();

    const added = grid.sheet(newGid);
    expect(grid.sheetTitles()).toEqual(["Payments"]);
    expect([added.rowCount, added.columnCount]).toEqual([
      topDataRowIndex + 1,
      2,
    ]);
    expect(added.tables).toEqual([
      {
        tableId: mintedTableId,
        name: "Payment Log",
        range: {
          startRowIndex: tableHeaderRowIndex,
          endRowIndex: topDataRowIndex + 1,
          startColumnIndex: 0,
          endColumnIndex: 2,
        },
        columnProperties: [
          { columnName: "Payee" },
          { columnIndex: 1, columnName: "Amount" },
        ],
      },
    ]);
  });

  it("writes the minted column IDs in the column ID row and leaves the other head rows and the body row empty", () => {
    const { grid, created } = createdOnNewTab();

    const values = grid.sheet(newGid).values({
      startRowIndex: 0,
      endRowIndex: topDataRowIndex + 1,
      startColumnIndex: 0,
      endColumnIndex: 2,
    });
    const filledCells = values.map((row) =>
      row.filter((cell) => cell !== null && cell !== ""),
    );
    const expected = Array.from(
      { length: topDataRowIndex + 1 },
      (): unknown[] => [],
    );
    expected[colIdRowIndex] = created.columnIds;
    expected[tableHeaderRowIndex] = ["Payee", "Amount"];
    expect(filledCells).toEqual(expected);
  });

  it("returns the Table's minted ID, a prefix from its name and a column ID per column under that prefix", () => {
    const { grid, created } = createdOnNewTab();

    expect(created).toEqual({
      tableId: grid.sheet(newGid).tables[0]?.tableId,
      idPrefix: "plg",
      columnIds: [mintedColumnId("plg"), mintedColumnId("plg")],
    });
    expect(created.columnIds[0]).not.toBe(created.columnIds[1]);
  });

  it("skips a prefix in use or retired", () => {
    stubSheetsService();
    const raw = SpreadsheetRaw.init();

    const created = raw.createTable({
      name: "Payment Log",
      placement: { kind: "newTab", sheetGid: newGid, title: "Payments" },
      idPrefix: { prefixesInUse: new Set(["plg"]), retiredPrefixes: ["plgg"] },
      columns: [{ header: "Payee" }],
    });

    expect(created.idPrefix).not.toMatch(/^plgg?$/);
  });

  it("sends a column's type and dropdown rule in the creation's column properties", () => {
    const { grid } = stubSheetsService();
    const raw = SpreadsheetRaw.init();

    raw.createTable({
      name: "Payment Log",
      placement: { kind: "newTab", sheetGid: newGid, title: "Payments" },
      idPrefix: noTakenIdPrefixes,
      columns: [
        { header: "Amount", columnType: { type: "CURRENCY", rule: null } },
        {
          header: "Status",
          columnType: {
            type: "DROPDOWN",
            rule: { condition: "ONE_OF_LIST", values: ["Paid", "Due"] },
          },
        },
      ],
    });
    raw.batchUpdateGSheets();

    expect(grid.sheet(newGid).tables[0]?.columnProperties).toEqual([
      { columnName: "Amount", columnType: "CURRENCY" },
      {
        columnIndex: 1,
        columnName: "Status",
        columnType: "DROPDOWN",
        dataValidationRule: {
          condition: {
            type: "ONE_OF_LIST",
            values: [{ userEnteredValue: "Paid" }, { userEnteredValue: "Due" }],
          },
        },
      },
    ]);
  });

  it("keeps a fixed Table ID, prefix and column IDs, and writes seed rows as its body", () => {
    const { grid } = stubSheetsService();
    const raw = SpreadsheetRaw.init();

    const created = raw.createTable({
      name: "Payment Log",
      placement: { kind: "newTab", sheetGid: newGid, title: "Payments" },
      tableId: "tbl-fixture01",
      idPrefix: "pay",
      columns: [
        { header: "Payee", columnId: "c:pay:payee" },
        { header: "Amount", columnId: "c:pay:amount" },
      ],
      seedRows: [
        ["Ann", 5],
        ["Bo", 7],
      ],
    });
    raw.batchUpdateGSheets();

    expect(created).toEqual({
      tableId: "tbl-fixture01",
      idPrefix: "pay",
      columnIds: ["c:pay:payee", "c:pay:amount"],
    });
    const added = grid.sheet(newGid);
    expect(added.tables[0]?.tableId).toBe("tbl-fixture01");
    expect(added.tables[0]?.range.endRowIndex).toBe(topDataRowIndex + 2);
    expect(
      added.values({
        startRowIndex: colIdRowIndex,
        endRowIndex: colIdRowIndex + 1,
      }),
    ).toEqual([["c:pay:payee", "c:pay:amount"]]);
    expect(added.values({ startRowIndex: topDataRowIndex })).toEqual([
      ["Ann", 5],
      ["Bo", 7],
    ]);
  });

  it("refuses a seed row wider than the columns, queuing nothing", () => {
    stubSheetsService();
    const raw = SpreadsheetRaw.init();

    expect(() =>
      raw.createTable({
        name: "Payment Log",
        placement: { kind: "newTab", sheetGid: newGid, title: "Payments" },
        idPrefix: "pay",
        columns: [{ header: "Payee" }],
        seedRows: [["Ann", 5]],
      }),
    ).toThrow(
      'Table "Payment Log" can\'t take seed row 1: it holds 2 values for 1 columns.',
    );
    expect(raw.writeOperations.addSheet).toEqual([]);
    expect(raw.writeOperations.addTable).toEqual([]);
  });
});

describe("SpreadsheetRaw.createTable beside the Tables on an existing tab", () => {
  const leftEnd = 3;
  const besideEnd = 6;

  function stubTwoTables() {
    return stubSheetsService({
      sheets: [
        {
          sheetId: existingGid,
          title: "Records",
          rows: buildGridRows({
            [tableHeaderRowIndex]: ["ID", "Name", "Qty", "", "Code"],
            [topDataRowIndex]: ["r1", "a", 1, "", "c1"],
          }),
          tables: [
            {
              tableId: "left",
              startColumnIndex: 0,
              endColumnIndex: leftEnd,
              endRowIndex: topDataRowIndex + 1,
            },
            {
              tableId: "beside",
              startColumnIndex: leftEnd + 1,
              endColumnIndex: besideEnd,
              endRowIndex: topDataRowIndex + 1,
            },
          ],
        },
      ],
    });
  }

  function createdBeside(gap?: 0 | 1) {
    const fake = stubTwoTables();
    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    const created = raw.createTable({
      name: "Payment Log",
      placement: {
        kind: "besideTables",
        sheetGid: existingGid,
        ...(gap === undefined ? {} : { gap }),
      },
      idPrefix: noTakenIdPrefixes,
      columns: [{ header: "Payee" }, { header: "Amount" }],
    });
    raw.batchUpdateGSheets();
    return { ...fake, created };
  }

  function createdRange(gap?: 0 | 1) {
    const { grid, created } = createdBeside(gap);
    const table = grid
      .sheet(existingGid)
      .tables.find(({ tableId }) => tableId === created.tableId);
    return table?.range;
  }

  it("places the Table right against the rightmost Table with no gap", () => {
    expect(createdRange()).toEqual({
      startRowIndex: tableHeaderRowIndex,
      endRowIndex: topDataRowIndex + 1,
      startColumnIndex: besideEnd,
      endColumnIndex: besideEnd + 2,
    });
    expect(createdRange(0)?.startColumnIndex).toBe(besideEnd);
  });

  it("leaves one empty column before it with a gap of 1", () => {
    expect(createdRange(1)).toEqual({
      startRowIndex: tableHeaderRowIndex,
      endRowIndex: topDataRowIndex + 1,
      startColumnIndex: besideEnd + 1,
      endColumnIndex: besideEnd + 3,
    });
  });

  it("grows the grid to hold it and leaves the other Tables and the cells between untouched", () => {
    const { grid } = createdBeside(1);

    const sheet = grid.sheet(existingGid);
    expect(sheet.columnCount).toBe(besideEnd + 3);
    expect(sheet.tables.map(({ tableId }) => tableId)).toEqual([
      "left",
      "beside",
      expect.any(String),
    ]);
    expect(
      sheet.values({
        startRowIndex: tableHeaderRowIndex,
        endRowIndex: topDataRowIndex + 1,
        endColumnIndex: besideEnd + 3,
      }),
    ).toEqual([
      ["ID", "Name", "Qty", "", "Code", null, null, "Payee", "Amount"],
      ["r1", "a", 1, "", "c1", null, null, null, null],
    ]);
  });

  it("writes its column IDs in its own columns of the column ID row", () => {
    const { grid, created } = createdBeside(1);

    expect(
      grid.sheet(existingGid).values({
        startRowIndex: colIdRowIndex,
        endRowIndex: colIdRowIndex + 1,
        startColumnIndex: besideEnd + 1,
        endColumnIndex: besideEnd + 3,
      }),
    ).toEqual([created.columnIds]);
  });

  it("fetches every tab's properties first when none are fetched, so it measures every Table on the tab", () => {
    const { grid } = stubTwoTables();
    const raw = SpreadsheetRaw.init();

    const created = raw.createTable({
      name: "Payment Log",
      placement: { kind: "besideTables", sheetGid: existingGid },
      idPrefix: "pay",
      columns: [{ header: "Payee" }],
    });
    raw.batchUpdateGSheets();

    expect(
      grid
        .sheet(existingGid)
        .tables.find(({ tableId }) => tableId === created.tableId)?.range
        .startColumnIndex,
    ).toBe(besideEnd);
  });

  it("writes seed rows inside its own columns", () => {
    const { grid } = stubTwoTables();
    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();

    raw.createTable({
      name: "Payment Log",
      placement: { kind: "besideTables", sheetGid: existingGid },
      idPrefix: "pay",
      columns: [
        { header: "Payee", columnId: "c:pay:payee" },
        { header: "Amount", columnId: "c:pay:amount" },
      ],
      seedRows: [["Ann", 5]],
    });
    raw.batchUpdateGSheets();

    expect(
      grid.sheet(existingGid).values({
        startRowIndex: topDataRowIndex,
        endRowIndex: topDataRowIndex + 1,
        startColumnIndex: besideEnd,
        endColumnIndex: besideEnd + 2,
      }),
    ).toEqual([["Ann", 5]]);
  });

  it("refuses while a Table on the tab has a column insert queued, since the insert would move the place it measured", () => {
    stubTwoTables();
    const raw = SpreadsheetRaw.init();
    raw.fetchAllSheetProperties();
    raw.table("beside").appendColumn({ columnId: "c:x:new", header: "New" });

    expect(() =>
      raw.createTable({
        name: "Payment Log",
        placement: { kind: "besideTables", sheetGid: existingGid },
        idPrefix: "pay",
        columns: [{ header: "Payee" }],
      }),
    ).toThrow(
      'Table "Payment Log" can\'t be placed on "Records" (gid 111) while a Table there has column changes queued: flush them first.',
    );
  });
});
