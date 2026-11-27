import { describe, expect, it } from "vitest";

import { tableLayout } from "../01_SpreadsheetSchema/tableLayout.ts";
import { fakeSheetsHttpTransport } from "../testSupport/fakeSheetsHttpTransport.ts";
import {
  type FakeCell,
  type FakeSheetsService,
  stubSheetsService,
} from "../testSupport/fakeSheetsService.ts";
import {
  SendSnapshotRestorer,
  type SnapshotRestore,
} from "./SendSnapshotRestorer.ts";
import type { SnapshotColumn } from "./SendSnapshotTaker.ts";

const sheetGid = 7;
const leftTableId = `fake-table-${sheetGid}`;
const rightTableId = `fake-table-${sheetGid}-1`;
const headerRowIndex = tableLayout.headRowOffsets.columnId;
const topDataRowIndex = tableLayout.headerZoneDepth;

// A column of one of the two Tables: its column ID, header and one data value.
interface SideColumn {
  columnId: string;
  header: string;
  value: string;
}

function sideColumn(name: string): SideColumn {
  return { columnId: `c:${name}`, header: name, value: `${name} value` };
}

const leftColumns = ["L0", "L1", "L2"].map(sideColumn);
const rightColumns = ["R0", "R1", "R2", "R3"].map(sideColumn);

function headRow(columns: SideColumn[], key: keyof SideColumn): FakeCell[] {
  return columns.map((column) => column[key]);
}

// Two Tables on one header row, the left one's columns then the right one's, with the deleted column gone from its Table.
function stubSideBySide(deleted: SideColumn): FakeSheetsService {
  const left = leftColumns.filter((column) => column !== deleted);
  const right = rightColumns.filter((column) => column !== deleted);
  const columns = [...left, ...right];
  const endRowIndex = topDataRowIndex + 1;
  return stubSheetsService({
    sheets: [
      {
        sheetId: sheetGid,
        title: "Side By Side",
        rows: [
          headRow(columns, "columnId"),
          [],
          [],
          headRow(columns, "header"),
          headRow(columns, "value"),
        ],
        tables: [
          {
            startRowIndex: headerRowIndex,
            startColumnIndex: 0,
            endRowIndex,
            endColumnIndex: left.length,
          },
          {
            startRowIndex: headerRowIndex,
            startColumnIndex: left.length,
            endRowIndex,
            endColumnIndex: columns.length,
          },
        ],
      },
    ],
  });
}

function savedColumn(
  column: SideColumn,
  colIndex: number,
  tableId: string,
): SnapshotColumn {
  return {
    sheetGid,
    tabTitle: "Side By Side",
    table: column.header.startsWith("L") ? "Left" : "Right",
    tableId,
    header: column.header,
    colIndex,
    columnIdRowIndex: 0,
    idColIndex: undefined,
    columnId: column.columnId,
    values: [column.columnId, "", "", column.header, column.value],
  };
}

function restoreOver(
  fake: FakeSheetsService,
  column: SnapshotColumn,
): SnapshotRestore {
  return SendSnapshotRestorer.init({
    spreadsheetId: "spreadsheet-under-test",
    transport: fakeSheetsHttpTransport(fake),
    snapshot: { columns: [column], fingerprints: [], versionHistoryItems: [] },
    isForced: false,
  }).restore();
}

function tableSpans(fake: FakeSheetsService): [string, number, number][] {
  return fake.grid
    .sheet(sheetGid)
    .tables.map(({ tableId, range }) => [
      tableId ?? "",
      range?.startColumnIndex ?? 0,
      range?.endColumnIndex ?? 0,
    ]);
}

const allColumns = [...leftColumns, ...rightColumns];

describe("restore of a column deleted on the shared edge of two side-by-side Tables", () => {
  it("recreates the right Table's first column inside the right Table", () => {
    const deleted = rightColumns[0] ?? sideColumn("R0");
    const fake = stubSideBySide(deleted);

    restoreOver(fake, savedColumn(deleted, leftColumns.length, rightTableId));

    expect(tableSpans(fake)).toEqual([
      [leftTableId, 0, leftColumns.length],
      [rightTableId, leftColumns.length, allColumns.length],
    ]);
    const values = fake.grid.sheet(sheetGid).values();
    expect(values[0]).toEqual(headRow(allColumns, "columnId"));
    expect(values[headerRowIndex]).toEqual(headRow(allColumns, "header"));
  });

  it("recreates the left Table's last column inside the left Table", () => {
    const deleted = leftColumns.at(-1) ?? sideColumn("L2");
    const fake = stubSideBySide(deleted);

    restoreOver(
      fake,
      savedColumn(deleted, leftColumns.length - 1, leftTableId),
    );

    expect(tableSpans(fake)).toEqual([
      [leftTableId, 0, leftColumns.length],
      [rightTableId, leftColumns.length, allColumns.length],
    ]);
    const values = fake.grid.sheet(sheetGid).values();
    expect(values[0]).toEqual(headRow(allColumns, "columnId"));
    expect(values[headerRowIndex]).toEqual(headRow(allColumns, "header"));
  });

  it("lists a column whose Table's ID is gone as not restored, recreating nothing", () => {
    const deleted = rightColumns[0] ?? sideColumn("R0");
    const fake = stubSideBySide(deleted);
    const before = fake.grid.sheet(sheetGid).values();
    const column = savedColumn(deleted, leftColumns.length, "fake-table-gone");

    const restore = restoreOver(fake, column);

    expect(restore.kind === "restored" && restore.goneColumns).toEqual([
      column,
    ]);
    expect(fake.grid.sheet(sheetGid).values()).toEqual(before);
  });
});
