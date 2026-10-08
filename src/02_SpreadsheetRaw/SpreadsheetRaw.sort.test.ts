import { describe, expect, it } from "vitest";

import {
  buildGridRows,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { SpreadsheetRaw } from "./SpreadsheetRaw";
import {
  startTableColIndex,
  tableHeaderRowIndex,
  topDataRowIndex,
} from "./spreadsheetRawTestSupport";

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
