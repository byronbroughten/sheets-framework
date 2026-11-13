import { describe, expect, it } from "vitest";

import type {
  GoogleRequest,
  ModeledRequestVerb,
} from "../00_Source/GoogleSheets/GoogleSheetsAPI";
import { assertType, type IsExactly } from "../testSupport/typeAssertions";
import { FlushClassifier, type FlushClassifierProps } from "./FlushClassifier";
import type {
  FlushRead,
  HeldItem,
} from "./FlushClassifier/flushClassifierTypes";
import type { requestJudgements } from "./FlushClassifier/requestJudgements";

type GoogleCellData = GoogleAppsScript.Sheets.Schema.CellData;

const gid = 7;

// The Payments Table: header on sheet row 5 (index 4), body rows at indexes 5 to 7, columns A to D.
const payments = {
  headerRow: 4,
  bodyRows: [5, 6, 7],
  endRow: 8,
  cols: { name: 0, id: 1, amount: 2, total: 3 },
  endCol: 4,
} as const;

function text(stringValue: string): GoogleCellData {
  return { effectiveValue: { stringValue }, userEnteredValue: {} };
}

function number(numberValue: number): GoogleCellData {
  return { effectiveValue: { numberValue }, userEnteredValue: {} };
}

function formula(formulaValue: string, numberValue: number): GoogleCellData {
  return {
    effectiveValue: { numberValue },
    userEnteredValue: { formulaValue },
  };
}

const empty: GoogleCellData = {};

// Head rows above the header (indexes 0 to 3), then the header, then three body rows; row 7 is wholly empty, so Google trims it.
const paymentsRows: GoogleCellData[][] = [
  [
    text("c:pay:name"),
    text("c:pay:id"),
    text("c:pay:amount"),
    text("c:pay:total"),
  ],
  [empty, empty, empty, empty],
  [empty, empty, empty, empty],
  [empty, empty, empty, empty],
  [text("Name"), text("ID"), text("Amount"), text("Total")],
  [text("Rent"), text("pay-0001"), number(1200), formula("=C6*2", 2400)],
  [text("Deposit"), empty, empty, formula("=C7*2", 0)],
];

function paymentsRead(rows: GoogleCellData[][] = paymentsRows): FlushRead {
  return {
    gridRanges: [
      {
        sheetId: gid,
        startRowIndex: 0,
        endRowIndex: payments.endRow,
        startColumnIndex: 0,
        endColumnIndex: payments.endCol,
      },
    ],
    spreadsheet: {
      sheets: [
        {
          properties: {
            sheetId: gid,
            title: "Payments",
            gridProperties: { rowCount: 20, columnCount: 6 },
          },
          tables: [
            {
              tableId: "tbl-pay",
              name: "Payments",
              range: {
                sheetId: gid,
                startRowIndex: payments.headerRow,
                endRowIndex: payments.endRow,
                startColumnIndex: 0,
                endColumnIndex: payments.endCol,
              },
              columnProperties: [
                { columnIndex: 0, columnName: "Name" },
                { columnIndex: 1, columnName: "ID" },
                { columnIndex: 2, columnName: "Amount", columnType: "DOUBLE" },
                { columnIndex: 3, columnName: "Total", columnType: "DOUBLE" },
              ],
            },
          ],
          data: [
            {
              startRow: 0,
              startColumn: 0,
              rowData: rows.map((values) => ({ values })),
            },
          ],
        },
      ],
    },
  };
}

function heldItems(
  requests: GoogleRequest[],
  props: Partial<FlushClassifierProps> = {},
): HeldItem[] {
  return FlushClassifier.init({
    reads: [paymentsRead()],
    requests,
    dropdownStyleReset: "ask",
    forcedTraits: [],
    ...props,
  }).heldItems();
}

function cellRange(rowIndex: number, colIndex: number) {
  return {
    sheetId: gid,
    startRowIndex: rowIndex,
    endRowIndex: rowIndex + 1,
    startColumnIndex: colIndex,
    endColumnIndex: colIndex + 1,
  };
}

function writeCell(
  rowIndex: number,
  colIndex: number,
  userEnteredValue: GoogleAppsScript.Sheets.Schema.ExtendedValue,
  fields = "userEnteredValue",
): GoogleRequest {
  return {
    updateCells: {
      range: cellRange(rowIndex, colIndex),
      rows: [{ values: [{ userEnteredValue }] }],
      fields,
    },
  };
}

// The entered-data items, each as its request, column, row and change.
function cellChanges(items: HeldItem[]) {
  return items.flatMap((item) => {
    if (item.kind !== "enteredData") return [];
    const { requestIndex, cell, change } = item;
    return [{ requestIndex, column: cell.column, row: cell.row, change }];
  });
}

const [rentRow, depositRow, emptyRow] = payments.bodyRows;
const { name, id, amount, total } = payments.cols;
const rentId = { rowId: "pay-0001" };

function rowRange(startRowIndex: number, endRowIndex: number) {
  return {
    sheetId: gid,
    startRowIndex,
    endRowIndex,
    startColumnIndex: 0,
    endColumnIndex: payments.endCol,
  };
}

function updateTable(
  tableId: string,
  columnType: string,
  fields = "columnProperties",
): GoogleRequest {
  return {
    updateTable: {
      table: {
        tableId,
        columnProperties: [
          { columnIndex: 2, columnName: "Amount", columnType },
        ],
      },
      fields,
    },
  };
}

describe("FlushClassifier.heldItems", () => {
  describe("entered data", () => {
    it("holds an overwrite of an entered value, by Table, column and row ID, before → after", () => {
      expect(
        heldItems([writeCell(rentRow, amount, { numberValue: 900 })]),
      ).toEqual([
        {
          kind: "enteredData",
          requestIndex: 0,
          verb: "updateCells",
          cell: {
            sheetGid: gid,
            table: "Payments",
            column: "Amount",
            row: rentId,
          },
          change: { kind: "overwrite", before: "1200", after: "900" },
        },
      ]);
    });

    it("holds a deleteRange of body rows per entered cell as deleted, and frees its formula", () => {
      const held = heldItems([
        {
          deleteRange: {
            range: rowRange(rentRow, rentRow + 1),
            shiftDimension: "ROWS",
          },
        },
      ]);
      expect(cellChanges(held)).toEqual([
        {
          requestIndex: 0,
          column: "Name",
          row: rentId,
          change: { kind: "delete", before: '"Rent"' },
        },
        {
          requestIndex: 0,
          column: "ID",
          row: rentId,
          change: { kind: "delete", before: '"pay-0001"' },
        },
        {
          requestIndex: 0,
          column: "Amount",
          row: rentId,
          change: { kind: "delete", before: "1200" },
        },
      ]);
    });

    it("holds a column deleteRange from the column ID row down per entered body cell as deleted", () => {
      const held = heldItems([
        {
          deleteRange: {
            range: {
              sheetId: gid,
              startRowIndex: 0,
              endRowIndex: payments.endRow,
              startColumnIndex: amount,
              endColumnIndex: amount + 1,
            },
            shiftDimension: "COLUMNS",
          },
        },
      ]);
      expect(cellChanges(held)).toEqual([
        {
          requestIndex: 0,
          column: "Amount",
          row: rentId,
          change: { kind: "delete", before: "1200" },
        },
      ]);
    });

    it("holds a deleteTable per entered cell of the read Table range as deleted, and frees its formulas", () => {
      const held = heldItems([{ deleteTable: { tableId: "tbl-pay" } }]);
      const deleted = (before: string) => ({ kind: "delete", before });
      const depositRowId = { sheetRow: depositRow + 1, nameValue: "Deposit" };
      expect(cellChanges(held)).toEqual([
        {
          requestIndex: 0,
          column: "Name",
          row: rentId,
          change: deleted('"Rent"'),
        },
        {
          requestIndex: 0,
          column: "ID",
          row: rentId,
          change: deleted('"pay-0001"'),
        },
        {
          requestIndex: 0,
          column: "Amount",
          row: rentId,
          change: deleted("1200"),
        },
        {
          requestIndex: 0,
          column: "Name",
          row: depositRowId,
          change: deleted('"Deposit"'),
        },
      ]);
    });

    it("frees a write after a deleteTable into a cell it blanked", () => {
      const held = heldItems([
        { deleteTable: { tableId: "tbl-pay" } },
        writeCell(rentRow, amount, { numberValue: 1 }),
      ]);
      expect(held.filter((item) => item.kind !== "enteredData")).toEqual([]);
      expect(
        cellChanges(held).map((change) => change.requestIndex),
      ).not.toContain(1);
    });

    it("holds a repeatCell over an entered value and frees the empty cells it also covers", () => {
      const held = heldItems([
        {
          repeatCell: {
            range: {
              sheetId: gid,
              startRowIndex: rentRow,
              endRowIndex: payments.endRow,
              startColumnIndex: amount,
              endColumnIndex: amount + 1,
            },
            cell: { userEnteredValue: { numberValue: 0 } },
            fields: "userEnteredValue",
          },
        },
      ]);
      expect(cellChanges(held)).toEqual([
        {
          requestIndex: 0,
          column: "Amount",
          row: rentId,
          change: { kind: "overwrite", before: "1200", after: "0" },
        },
      ]);
    });

    it("holds a PASTE_FORMULA pasteData over a typed value", () => {
      const held = heldItems([
        {
          pasteData: {
            coordinate: {
              sheetId: gid,
              rowIndex: rentRow,
              columnIndex: amount,
            },
            data: "=1+1",
            delimiter: ",",
            type: "PASTE_FORMULA",
          },
        },
      ]);
      expect(cellChanges(held)).toEqual([
        {
          requestIndex: 0,
          column: "Amount",
          row: rentId,
          change: { kind: "overwrite", before: "1200", after: "=1+1" },
        },
      ]);
    });

    it("names a row with no ID by its sheet row and Name value", () => {
      const held = heldItems([
        writeCell(depositRow, name, { stringValue: "Bond" }),
      ]);
      expect(cellChanges(held)).toEqual([
        {
          requestIndex: 0,
          column: "Name",
          row: { sheetRow: depositRow + 1, nameValue: "Deposit" },
          change: { kind: "overwrite", before: '"Deposit"', after: '"Bond"' },
        },
      ]);
    });

    it("holds a cutPaste landing on entered data outside its own source", () => {
      const held = heldItems([
        {
          cutPaste: {
            source: cellRange(rentRow, amount),
            destination: {
              sheetId: gid,
              rowIndex: depositRow,
              columnIndex: name,
            },
            pasteType: "PASTE_NORMAL",
          },
        },
      ]);
      expect(cellChanges(held)).toEqual([
        {
          requestIndex: 0,
          column: "Name",
          row: { sheetRow: depositRow + 1, nameValue: "Deposit" },
          change: { kind: "overwrite", before: '"Deposit"', after: "1200" },
        },
      ]);
    });

    it("holds a cutPaste that pastes no values as emptying its source", () => {
      const held = heldItems([
        {
          cutPaste: {
            source: cellRange(rentRow, amount),
            destination: {
              sheetId: gid,
              rowIndex: emptyRow,
              columnIndex: amount,
            },
            pasteType: "PASTE_FORMAT",
          },
        },
      ]);
      expect(cellChanges(held)).toEqual([
        {
          requestIndex: 0,
          column: "Amount",
          row: rentId,
          change: { kind: "overwrite", before: "1200", after: "(empty)" },
        },
      ]);
    });
  });

  describe("other held kinds", () => {
    it("holds every findReplace", () => {
      expect(
        heldItems([
          {
            findReplace: {
              find: "Rent",
              replacement: "Lease",
              allSheets: true,
            },
          },
        ]),
      ).toEqual([
        {
          kind: "findReplace",
          requestIndex: 0,
          find: "Rent",
          replacement: "Lease",
        },
      ]);
    });

    it("holds a dropdown reset under ask", () => {
      expect(heldItems([updateTable("tbl-pay", "DROPDOWN")])).toEqual([
        { kind: "dropdownReset", requestIndex: 0, table: "Payments" },
      ]);
    });

    it("frees a dropdown reset under allow", () => {
      expect(
        heldItems([updateTable("tbl-pay", "DROPDOWN")], {
          dropdownStyleReset: "allow",
        }),
      ).toEqual([]);
    });

    it("frees a column-properties update on a Table with no dropdown, as read or as sent", () => {
      expect(heldItems([updateTable("tbl-pay", "DOUBLE")])).toEqual([]);
    });

    it("frees a dropdown reset on a Table the same flush added", () => {
      expect(
        heldItems([
          {
            addTable: {
              table: {
                tableId: "tbl-new",
                name: "New",
                range: rowRange(10, 12),
              },
            },
          },
          updateTable("tbl-new", "DROPDOWN"),
        ]),
      ).toEqual([]);
    });

    it("holds each forced trait ahead of the requests", () => {
      const trait = { table: "Payments", column: "Amount", trait: "currency" };
      expect(
        heldItems([writeCell(rentRow, amount, { numberValue: 1 })], {
          forcedTraits: [trait],
        })[0],
      ).toEqual({ kind: "forcedTrait", trait });
    });

    it("holds a request verb with no classification", () => {
      expect(heldItems([{ deleteSheet: { sheetId: gid } }])).toEqual([
        { kind: "unclassified", requestIndex: 0, verb: "deleteSheet" },
      ]);
    });

    it("holds a cell on a sheet no read covered", () => {
      expect(
        heldItems([
          {
            updateCells: {
              range: { ...cellRange(rentRow, amount), sheetId: 99 },
              rows: [{ values: [{ userEnteredValue: { numberValue: 1 } }] }],
              fields: "userEnteredValue",
            },
          },
        ]),
      ).toEqual([
        {
          kind: "unreadCell",
          requestIndex: 0,
          verb: "updateCells",
          cell: {
            sheetGid: 99,
            table: undefined,
            column: "C",
            row: { sheetRow: rentRow + 1, nameValue: undefined },
          },
        },
      ]);
    });

    it("holds a Table cell outside every range a read asked for", () => {
      const read = paymentsRead();
      const held = heldItems(
        [writeCell(emptyRow, amount, { numberValue: 1 })],
        {
          reads: [{ ...read, gridRanges: [rowRange(0, emptyRow)] }],
        },
      );
      expect(held).toEqual([
        {
          kind: "unreadCell",
          requestIndex: 0,
          verb: "updateCells",
          cell: {
            sheetGid: gid,
            table: "Payments",
            column: "Amount",
            row: { sheetRow: emptyRow + 1, nameValue: undefined },
          },
        },
      ]);
    });

    it("holds a deleteTable of a Table no read fetched as unmeasured", () => {
      expect(heldItems([{ deleteTable: { tableId: "tbl-gone" } }])).toEqual([
        expect.objectContaining({
          kind: "unmeasured",
          requestIndex: 0,
          verb: "deleteTable",
        }),
      ]);
    });

    it("frees a deleteTable of a Table the same flush added", () => {
      expect(
        heldItems([
          {
            addTable: {
              table: {
                tableId: "tbl-new",
                name: "New",
                range: rowRange(10, 12),
              },
            },
          },
          { deleteTable: { tableId: "tbl-new" } },
        ]),
      ).toEqual([]);
    });

    it("holds a write into a range an earlier sort reordered as unmeasured", () => {
      const held = heldItems([
        {
          sortRange: {
            range: rowRange(rentRow, payments.endRow),
            sortSpecs: [{ dimensionIndex: amount, sortOrder: "ASCENDING" }],
          },
        },
        writeCell(depositRow, amount, { numberValue: 1 }),
      ]);
      expect(held).toEqual([
        expect.objectContaining({
          kind: "unmeasured",
          requestIndex: 1,
          verb: "updateCells",
        }),
      ]);
    });

    it("holds an open-ended range on a tab no read sized as unmeasured", () => {
      expect(
        heldItems([
          {
            updateCells: {
              range: { sheetId: 99, startRowIndex: 0 },
              fields: "userEnteredValue",
            },
          },
        ]),
      ).toEqual([
        {
          kind: "unmeasured",
          requestIndex: 0,
          verb: "updateCells",
          detail: "an open-ended range on gid 99, a tab no read sized",
        },
      ]);
    });

    it("holds a tab's grid properties as unmeasured and frees its title", () => {
      expect(
        heldItems([
          {
            updateSheetProperties: {
              properties: { sheetId: gid, gridProperties: { rowCount: 5 } },
              fields: "gridProperties.rowCount",
            },
          },
          {
            updateSheetProperties: {
              properties: { sheetId: gid, title: "Paid" },
              fields: "title",
            },
          },
        ]),
      ).toEqual([
        {
          kind: "unmeasured",
          requestIndex: 0,
          verb: "updateSheetProperties",
          detail: "a tab's grid properties, which can drop rows or columns",
        },
      ]);
    });
  });

  describe("free", () => {
    it("frees writes into empty, formula and trimmed cells", () => {
      expect(
        heldItems([
          writeCell(depositRow, amount, { numberValue: 5 }),
          writeCell(rentRow, total, { numberValue: 5 }),
          writeCell(emptyRow, amount, { numberValue: 5 }),
        ]),
      ).toEqual([]);
    });

    it("frees a write equal to what the cell holds", () => {
      expect(
        heldItems([writeCell(rentRow, amount, { numberValue: 1200 })]),
      ).toEqual([]);
    });

    it("frees a mask that leaves userEnteredValue alone", () => {
      expect(
        heldItems([
          writeCell(rentRow, amount, { numberValue: 9 }, "userEnteredFormat"),
        ]),
      ).toEqual([]);
    });

    it("frees head rows and the header row", () => {
      expect(
        heldItems([
          writeCell(0, name, { stringValue: "c:pay:label" }),
          writeCell(payments.headerRow, amount, { stringValue: "Sum" }),
        ]),
      ).toEqual([]);
    });

    it("frees appends", () => {
      expect(
        heldItems([
          {
            appendCells: {
              sheetId: gid,
              rows: [
                { values: [{ userEnteredValue: { stringValue: "Fee" } }] },
              ],
              fields: "userEnteredValue",
            },
          },
          { appendDimension: { sheetId: gid, dimension: "ROWS", length: 5 } },
        ]),
      ).toEqual([]);
    });

    it("frees a write into rows or cells the same flush inserted", () => {
      expect(
        heldItems([
          {
            insertDimension: {
              range: {
                sheetId: gid,
                dimension: "ROWS",
                startIndex: rentRow,
                endIndex: rentRow + 1,
              },
            },
          },
          writeCell(rentRow, amount, { numberValue: 1 }),
          {
            insertRange: {
              range: cellRange(rentRow, id),
              shiftDimension: "ROWS",
            },
          },
          writeCell(rentRow, id, { stringValue: "pay-0002" }),
        ]),
      ).toEqual([]);
    });

    it("frees a moveDimension and a cutPaste into a gap", () => {
      expect(
        heldItems([
          {
            moveDimension: {
              source: {
                sheetId: gid,
                dimension: "COLUMNS",
                startIndex: amount,
                endIndex: amount + 1,
              },
              destinationIndex: payments.endCol,
            },
          },
          {
            cutPaste: {
              source: cellRange(rentRow, name),
              destination: {
                sheetId: gid,
                rowIndex: emptyRow,
                columnIndex: name,
              },
              pasteType: "PASTE_NORMAL",
            },
          },
        ]),
      ).toEqual([]);
    });
  });

  describe("shift tracking", () => {
    it("maps a write below an insertDimension back to the shifted read cell", () => {
      const held = heldItems([
        {
          insertDimension: {
            range: {
              sheetId: gid,
              dimension: "ROWS",
              startIndex: rentRow,
              endIndex: rentRow + 2,
            },
          },
        },
        writeCell(rentRow + 2, amount, { numberValue: 1 }),
      ]);
      expect(cellChanges(held)).toEqual([
        {
          requestIndex: 1,
          column: "Amount",
          row: rentId,
          change: { kind: "overwrite", before: "1200", after: "1" },
        },
      ]);
    });

    it("maps a write after a deleteDimension back to the row that moved up", () => {
      const held = heldItems([
        {
          deleteDimension: {
            range: {
              sheetId: gid,
              dimension: "ROWS",
              startIndex: rentRow,
              endIndex: rentRow + 1,
            },
          },
        },
        writeCell(rentRow, name, { stringValue: "Bond" }),
      ]);
      expect(cellChanges(held).filter((c) => c.requestIndex === 1)).toEqual([
        {
          requestIndex: 1,
          column: "Name",
          row: { sheetRow: depositRow + 1, nameValue: "Deposit" },
          change: { kind: "overwrite", before: '"Deposit"', after: '"Bond"' },
        },
      ]);
    });

    it("maps a write after a moveDimension back to the moved row", () => {
      const held = heldItems([
        {
          moveDimension: {
            source: {
              sheetId: gid,
              dimension: "ROWS",
              startIndex: rentRow,
              endIndex: rentRow + 1,
            },
            destinationIndex: depositRow + 1,
          },
        },
        writeCell(rentRow, amount, { numberValue: 1 }),
        writeCell(depositRow, amount, { numberValue: 1 }),
      ]);
      expect(cellChanges(held)).toEqual([
        {
          requestIndex: 2,
          column: "Amount",
          row: rentId,
          change: { kind: "overwrite", before: "1200", after: "1" },
        },
      ]);
    });

    it("maps a write before a backward moveDimension's source back to the row moved down", () => {
      const held = heldItems([
        {
          moveDimension: {
            source: {
              sheetId: gid,
              dimension: "ROWS",
              startIndex: depositRow,
              endIndex: depositRow + 1,
            },
            destinationIndex: rentRow,
          },
        },
        writeCell(rentRow, amount, { numberValue: 1 }),
        writeCell(depositRow, amount, { numberValue: 1 }),
      ]);
      expect(cellChanges(held)).toEqual([
        {
          requestIndex: 2,
          column: "Amount",
          row: rentId,
          change: { kind: "overwrite", before: "1200", after: "1" },
        },
      ]);
    });

    it("maps a write right of an inserted column back to the column that moved over", () => {
      const held = heldItems([
        {
          insertDimension: {
            range: {
              sheetId: gid,
              dimension: "COLUMNS",
              startIndex: name,
              endIndex: name + 1,
            },
          },
        },
        writeCell(rentRow, name, { stringValue: "Lease" }),
        writeCell(rentRow, amount + 1, { numberValue: 1 }),
      ]);
      expect(cellChanges(held)).toEqual([
        {
          requestIndex: 2,
          column: "Amount",
          row: rentId,
          change: { kind: "overwrite", before: "1200", after: "1" },
        },
      ]);
    });

    it("maps a write after a cutPaste back to the cut cell, and frees its emptied source", () => {
      const held = heldItems([
        {
          cutPaste: {
            source: cellRange(rentRow, amount),
            destination: {
              sheetId: gid,
              rowIndex: emptyRow,
              columnIndex: amount,
            },
            pasteType: "PASTE_NORMAL",
          },
        },
        writeCell(emptyRow, amount, { numberValue: 1 }),
        writeCell(rentRow, amount, { numberValue: 1 }),
      ]);
      expect(cellChanges(held)).toEqual([
        {
          requestIndex: 1,
          column: "Amount",
          row: rentId,
          change: { kind: "overwrite", before: "1200", after: "1" },
        },
      ]);
    });

    it("shifts only the band a range insert covers", () => {
      const held = heldItems([
        {
          insertRange: {
            range: cellRange(rentRow, amount),
            shiftDimension: "ROWS",
          },
        },
        writeCell(depositRow, amount, { numberValue: 1 }),
        writeCell(depositRow, id, { stringValue: "pay-0002" }),
        writeCell(rentRow, id, { stringValue: "pay-0003" }),
      ]);
      expect(cellChanges(held)).toEqual([
        {
          requestIndex: 1,
          column: "Amount",
          row: rentId,
          change: { kind: "overwrite", before: "1200", after: "1" },
        },
        {
          requestIndex: 3,
          column: "ID",
          row: rentId,
          change: {
            kind: "overwrite",
            before: '"pay-0001"',
            after: '"pay-0003"',
          },
        },
      ]);
    });
  });

  it("classifies every ModeledRequestVerb, so a new one can't go unjudged", () => {
    assertType<
      IsExactly<
        Exclude<ModeledRequestVerb, keyof typeof requestJudgements>,
        never
      >
    >(true);
  });
});
