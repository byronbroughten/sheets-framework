import { Val } from "@byronbroughten/utils/val";
import { describe, expect, it, vi } from "vitest";

import {
  googleRawRequest,
  type GoogleRequest,
  type GoogleSheetsAPI,
  type SheetsHttpRequest,
} from "../00_Source/GoogleSheets/GoogleSheetsAPI";
import { HeldFlushRefusal } from "../00_Source/RawSource/HeldFlushRefusal";
import { SheetIndex } from "../00_Source/RawSource/SheetIndex";
import { getColumnTraitByName } from "../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import { installedConfigs } from "../01_SpreadsheetSchema/configReaders/configRegister";
import { getTableTraitByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { SpreadsheetRaw } from "../02_SpreadsheetRaw/SpreadsheetRaw";
import type { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import type { Endpoints } from "../06_API/Endpoints";
import { expectedSheetLayout } from "../testSupport/expectedSheetLayout";
import {
  buildGridRows,
  type FakeCell,
  type FakeCellValue,
} from "../testSupport/fakeSheetsService";
import type { HeldList } from "./EnteredDataGate";
import { NodeHost } from "./NodeHost";

type GoogleCellData = GoogleAppsScript.Sheets.Schema.CellData;

const spreadsheetId = "spreadsheet-under-test";
const gadgetsGid = 111;
const gadgetsTableId = "fake-table";

const gadgetsPayload = {
  sheets: [
    {
      properties: { sheetId: gadgetsGid, title: "Gadgets" },
      tables: [
        {
          tableId: gadgetsTableId,
          range: {
            startRowIndex: expectedSheetLayout.tableHeaderRowIndex,
            endRowIndex: 11,
            startColumnIndex: expectedSheetLayout.startTableColIndex,
            endColumnIndex: 5,
          },
        },
      ],
    },
  ],
};

function seedHost(isDryRun: boolean) {
  const transport = vi.fn((_request: SheetsHttpRequest) => gadgetsPayload);
  const host = NodeHost.init({
    configs: installedConfigs(),
    spreadsheetId,
    transport,
    isDryRun,
    log: vi.fn(),
  }).ensureGlobals();
  return { host, transport };
}

function writeOneCell(): SpreadsheetRaw {
  const raw = SpreadsheetRaw.init();
  raw.fetchAllSheetProperties();
  raw.table(gadgetsTableId).row(1).cell(2).updateValue("Processing...");
  raw.batchUpdateGSheets();
  return raw;
}

describe("NodeHost.ensureGlobals", () => {
  it("runs the framework's read against the live spreadsheet while a dry run is armed", () => {
    const { transport } = seedHost(true);

    SpreadsheetRaw.init().fetchAllSheetProperties();

    expect(transport).toHaveBeenCalledTimes(1);
    expect(transport.mock.calls[0]?.[0].method).toBe("GET");
  });

  it("puts no write on the wire while a dry run is armed, whoever calls the flush", () => {
    const { host, transport } = seedHost(true);

    writeOneCell();

    expect(
      transport.mock.calls.filter(([request]) => request.method === "POST"),
    ).toEqual([]);
    expect(host.summary.count).toBe(1);
    expect(host.summary.lines[0]).toContain(`gid ${gadgetsGid}!C6:C6`);
  });

  it("sends the write once the dry run is not armed", () => {
    const { host, transport } = seedHost(false);

    writeOneCell();

    const posts = transport.mock.calls.filter(
      ([request]) => request.method === "POST",
    );
    expect(posts).toHaveLength(1);
    expect(posts[0]?.[0].url).toContain(":batchUpdate");
    expect(host.summary.count).toBe(1);
  });

  it("reads the spreadsheet it was given", () => {
    const { transport } = seedHost(true);

    SpreadsheetRaw.init().fetchAllSheetProperties();

    expect(transport.mock.calls[0]?.[0].url).toContain(`/${spreadsheetId}?`);
  });

  it("installs Logger only, not a Sheets or PropertiesService global", () => {
    seedHost(true);

    const globals = globalThis as {
      Sheets?: unknown;
      PropertiesService?: unknown;
    };
    expect(globals.Sheets).toBeUndefined();
    expect(globals.PropertiesService).toBeUndefined();
  });
});

// The Payments Table on its own sheet: header at index 0, Rent at 1, an empty row at 2, a Total formula at 3, columns A to C.
const payments = {
  gid: 7,
  rentRow: 1,
  emptyRow: 2,
  totalRow: 3,
  endRow: 4,
  endCol: 3,
};
const paymentsAmountCol = 2;
const gateTarget = "app:spreadsheet-under-test";

function paymentsSheet(rows?: GoogleCellData[][]) {
  return {
    properties: {
      sheetId: payments.gid,
      title: "Payments",
      gridProperties: { rowCount: 20, columnCount: 6 },
    },
    tables: [
      {
        tableId: "tbl-pay",
        name: "Payments",
        range: {
          sheetId: payments.gid,
          startRowIndex: 0,
          endRowIndex: payments.endRow,
          startColumnIndex: 0,
          endColumnIndex: payments.endCol,
        },
        columnProperties: [
          { columnIndex: 0, columnName: "Name" },
          { columnIndex: 1, columnName: "ID" },
          { columnIndex: 2, columnName: "Amount", columnType: "DOUBLE" },
        ],
      },
    ],
    data:
      rows === undefined
        ? undefined
        : [
            {
              startRow: 0,
              startColumn: 0,
              rowData: rows.map((values) => ({ values })),
            },
          ],
  };
}

// The Node host's mask gives a non-formula cell an empty userEnteredValue.
function text(stringValue: string): GoogleCellData {
  return { effectiveValue: { stringValue }, userEnteredValue: {} };
}

function amount(numberValue: number): GoogleCellData {
  return { effectiveValue: { numberValue }, userEnteredValue: {} };
}

const paymentsRows: GoogleCellData[][] = [
  [text("Name"), text("ID"), text("Amount")],
  [text("Rent"), text("pay-0001"), amount(1200)],
  [],
  [
    text("Total"),
    text("pay-0002"),
    {
      effectiveValue: { numberValue: 1200 },
      userEnteredValue: { formulaValue: "=SUM(C2:C3)" },
    },
  ],
];

// Properties alone on a get; the grid on a getByDataFilter; an empty reply to a batchUpdate.
function gateTransport() {
  return vi.fn((request: SheetsHttpRequest) => {
    if (request.url.includes(":batchUpdate")) return { replies: [] };
    if (request.url.includes(":getByDataFilter")) {
      return { sheets: [paymentsSheet(paymentsRows)] };
    }
    return { sheets: [paymentsSheet()] };
  });
}

interface GatedHostCase {
  isDryRun: boolean;
  target?: string;
  confirmedList?: HeldList;
}

interface SendCase extends Omit<GatedHostCase, "isDryRun"> {
  flushes: GoogleRequest[][];
}

function seedGatedHost({ isDryRun, target, confirmedList }: GatedHostCase) {
  const transport = gateTransport();
  const host = NodeHost.init({
    configs: installedConfigs(),
    spreadsheetId,
    transport,
    isDryRun,
    log: vi.fn(),
    enteredDataGate: {
      target: target ?? gateTarget,
      confirmedList,
      hash: (input) => `hash(${input})`,
    },
  });
  return { host, api: host.googleSheetsAPI, transport };
}

function readPayments(api: GoogleSheetsAPI): void {
  api.fetchGrid(
    [
      {
        sheetId: payments.gid,
        startRowIndex: SheetIndex.row(0),
        endRowIndex: SheetIndex.row(payments.endRow),
        startColumnIndex: SheetIndex.col(0),
        endColumnIndex: SheetIndex.col(payments.endCol),
      },
    ],
    { includeProgrammaticFacts: false },
  );
}

function flushRequests(api: GoogleSheetsAPI, requests: GoogleRequest[]): void {
  api.flush(
    requests.map((request) => ({
      kind: "raw",
      request: googleRawRequest(request),
    })),
  );
}

// Amounts down the Amount column, one row each from startRowIndex.
function writeAmounts(
  startRowIndex: number,
  numberValues: number[],
): GoogleRequest {
  return {
    updateCells: {
      range: {
        sheetId: payments.gid,
        startRowIndex,
        endRowIndex: startRowIndex + numberValues.length,
        startColumnIndex: paymentsAmountCol,
        endColumnIndex: paymentsAmountCol + 1,
      },
      rows: numberValues.map((numberValue) => ({
        values: [{ userEnteredValue: { numberValue } }],
      })),
      fields: "userEnteredValue",
    },
  };
}

const appendRows: GoogleRequest = {
  appendDimension: { sheetId: payments.gid, dimension: "ROWS", length: 5 },
};

const replaceRent: GoogleRequest = {
  findReplace: { find: "Rent", replacement: "Lease", allSheets: true },
};

function batchUpdatePosts(transport: ReturnType<typeof gateTransport>) {
  return transport.mock.calls.filter(([request]) =>
    request.url.includes(":batchUpdate"),
  );
}

function heldKinds(heldList: HeldList | undefined): string[][] | undefined {
  return heldList?.flushes.map((items) => items.map(({ kind }) => kind));
}

// A dry run of the given flushes after one read of the Payments grid, and the held list it leaves.
function dryRunHeldList(flushes: GoogleRequest[][], target?: string): HeldList {
  const { host, api } = seedGatedHost({ isDryRun: true, target });
  readPayments(api);
  for (const requests of flushes) flushRequests(api, requests);
  return Val.assert(host.heldList, "heldList");
}

function sendFlushes({ flushes, confirmedList, target }: SendCase) {
  const { host, api, transport } = seedGatedHost({
    isDryRun: false,
    target,
    confirmedList,
  });
  readPayments(api);
  let refusal: unknown;
  try {
    for (const requests of flushes) flushRequests(api, requests);
  } catch (error) {
    refusal = error;
  }
  return { host, refusal, posts: batchUpdatePosts(transport) };
}

describe("NodeHost with an entered-data gate", () => {
  it.each([
    { label: "frees an append", request: appendRows, kinds: [] },
    {
      label: "frees a write into an empty cell a read covered",
      request: writeAmounts(payments.emptyRow, [5]),
      kinds: [],
    },
    {
      label: "holds a findReplace",
      request: replaceRent,
      kinds: ["findReplace"],
    },
    {
      label: "holds an overwrite of entered data",
      request: writeAmounts(payments.rentRow, [900]),
      kinds: ["enteredData"],
    },
    {
      label: "frees an overwrite of a formula cell",
      request: writeAmounts(payments.totalRow, [900]),
      kinds: [],
    },
  ])("$label", ({ request, kinds }) => {
    expect(heldKinds(dryRunHeldList([[request]]))).toEqual([kinds]);
  });

  it("holds only the entered cell of a write spanning entered, empty and formula cells", () => {
    const heldList = dryRunHeldList([
      [writeAmounts(payments.rentRow, [900, 5, 900])],
    ]);

    expect(heldList.flushes).toEqual([
      [
        expect.objectContaining({
          kind: "enteredData",
          cell: expect.objectContaining({
            row: { rowId: "pay-0001" },
          }),
        }),
      ],
    ]);
  });

  it("asks each grid read for every cell's formula", () => {
    const { api, transport } = seedGatedHost({ isDryRun: true });

    readPayments(api);

    const [read] = transport.mock.calls.filter(([request]) =>
      request.url.includes(":getByDataFilter"),
    );
    expect(decodeURIComponent(read?.[0].url ?? "")).toContain(
      "rowData(values(effectiveValue,userEnteredValue(formulaValue)))",
    );
  });

  it("keeps a free flush's empty list, so each flush's items sit at its index", () => {
    expect(heldKinds(dryRunHeldList([[appendRows], [replaceRent]]))).toEqual([
      [],
      ["findReplace"],
    ]);
  });

  it("judges a flush against the reads since the one before it, as the send will", () => {
    expect(
      heldKinds(
        dryRunHeldList([[appendRows], [writeAmounts(payments.emptyRow, [5])]]),
      ),
    ).toEqual([[], ["unreadCell"]]);
  });

  it("holds a write into a cell no read covered", () => {
    const { host, api } = seedGatedHost({ isDryRun: true });
    api.fetchSheetProperties();

    flushRequests(api, [writeAmounts(payments.emptyRow, [5])]);

    expect(heldKinds(host.heldList)).toEqual([["unreadCell"]]);
  });

  it("withholds every flush on a dry run, free ones included, and still reports them", () => {
    const { host, api, transport } = seedGatedHost({ isDryRun: true });
    readPayments(api);

    flushRequests(api, [appendRows]);
    flushRequests(api, [writeAmounts(payments.rentRow, [900])]);

    expect(batchUpdatePosts(transport)).toEqual([]);
    expect(host.summary.count).toBe(2);
  });

  it("gives one token for one target and list, and another when either differs", () => {
    const flushes = [[writeAmounts(payments.rentRow, [900])]];
    const { token } = dryRunHeldList(flushes);

    expect(dryRunHeldList(flushes).token).toBe(token);
    expect(dryRunHeldList(flushes, "app:another-spreadsheet").token).not.toBe(
      token,
    );
    expect(
      dryRunHeldList([[writeAmounts(payments.rentRow, [800])]]).token,
    ).not.toBe(token);
  });

  it("sends a free flush with no token", () => {
    const { refusal, posts } = sendFlushes({ flushes: [[appendRows]] });

    expect(refusal).toBeUndefined();
    expect(posts).toHaveLength(1);
  });

  it("refuses a held flush as uncovered when no token is given, and sends nothing", () => {
    const { refusal, posts } = sendFlushes({
      flushes: [[writeAmounts(payments.rentRow, [900])]],
    });

    expect(refusal).toBeInstanceOf(HeldFlushRefusal);
    expect(refusal).toMatchObject({
      kind: "uncovered",
      landedFlushes: 0,
      flushCount: undefined,
    });
    expect(posts).toEqual([]);
  });

  it("sends a held flush the dry run's token and list cover", () => {
    const flushes = [[appendRows], [writeAmounts(payments.rentRow, [900])]];

    const { refusal, posts } = sendFlushes({
      flushes,
      confirmedList: dryRunHeldList(flushes),
    });

    expect(refusal).toBeUndefined();
    expect(posts).toHaveLength(2);
  });

  it("refuses as mismatched when the list's entry for the flush differs from its items, counting the list's flushes", () => {
    const covered = dryRunHeldList([
      [writeAmounts(payments.rentRow, [900])],
      [appendRows],
    ]);

    const { refusal, posts } = sendFlushes({
      flushes: [[writeAmounts(payments.rentRow, [800])]],
      confirmedList: covered,
    });

    expect(refusal).toMatchObject({
      kind: "mismatched",
      landedFlushes: 0,
      flushCount: 2,
    });
    expect(posts).toEqual([]);
  });

  it("refuses as mismatched when the list was edited after the token was made", () => {
    const flushes = [[writeAmounts(payments.rentRow, [800])]];
    const covered = dryRunHeldList([[writeAmounts(payments.rentRow, [900])]]);
    const edited = { ...covered, flushes: dryRunHeldList(flushes).flushes };

    const { refusal } = sendFlushes({ flushes, confirmedList: edited });

    expect(refusal).toMatchObject({ kind: "mismatched" });
  });

  it("refuses as mismatched when the token was made for another target", () => {
    const flushes = [[writeAmounts(payments.rentRow, [900])]];

    const { refusal } = sendFlushes({
      flushes,
      confirmedList: dryRunHeldList(flushes, "app:another-spreadsheet"),
    });

    expect(refusal).toMatchObject({ kind: "mismatched" });
  });

  it("refuses partway with the earlier flushes landed and counted", () => {
    const { refusal, posts } = sendFlushes({
      flushes: [
        [appendRows],
        [appendRows],
        [writeAmounts(payments.rentRow, [900])],
      ],
    });

    expect(refusal).toMatchObject({ kind: "uncovered", landedFlushes: 2 });
    expect(posts).toHaveLength(2);
  });

  it("classifies nothing without a gate, so a chore's write goes out", () => {
    const { host, transport } = seedHost(false);

    writeOneCell();

    expect(host.heldList).toBeUndefined();
    expect(
      transport.mock.calls.filter(([request]) => request.method === "POST"),
    ).toHaveLength(1);
  });
});

const topDataRowIndex = 4;

interface AppendCase {
  sheetGid: number;
  title: string;
  columnIds: string[];
  topRow: FakeCell[];
  append: (ss: SpreadsheetNamed) => void;
  endpoints?: Endpoints;
}

// A fresh program: the host installs its endpoints once, as a chore's would.
async function appendBeneathOneRow({
  sheetGid,
  title,
  columnIds,
  topRow,
  append,
  endpoints,
}: AppendCase): Promise<FakeCellValue[][]> {
  vi.resetModules();
  const { NodeHost } = await import("./NodeHost");
  const { stubSheetsService } =
    await import("../testSupport/fakeSheetsService");
  const { SpreadsheetNamed } =
    await import("../04_SpreadsheetNamed/SpreadsheetNamed");
  NodeHost.init({
    configs: installedConfigs(),
    spreadsheetId,
    transport: vi.fn(),
    isDryRun: false,
    log: vi.fn(),
    endpoints,
  }).ensureGlobals();
  const { grid } = stubSheetsService({
    sheets: [
      {
        sheetId: sheetGid,
        title,
        rows: buildGridRows({ 0: columnIds, [topDataRowIndex]: topRow }),
        table: { endRowIndex: topDataRowIndex + 1 },
      },
    ],
  });
  const ss = SpreadsheetNamed.init();
  append(ss);
  ss.batchUpdateGSheets();
  return grid.sheet(sheetGid).values({ startRowIndex: topDataRowIndex });
}

const runItemWithStatusOnly = {
  sheetGid: getTableTraitByName("runItem", "sheetGid"),
  title: "Run item",
  columnIds: [
    getColumnTraitByName("runItem", "id", "columnId"),
    getColumnTraitByName("runItem", "result", "columnId"),
    getColumnTraitByName("runItem", "runStatus", "columnId"),
  ],
  topRow: [null, null, "Succeeded"],
  append: (ss: SpreadsheetNamed) => {
    const sheet = ss.table("runItem");
    sheet.row(topDataRowIndex).prepFetchFull();
    ss.fetchAllPrepped();
    sheet.appendRowWithVals({ result: "appended" });
  },
};

describe("NodeHost.ensureGlobals, the endpoints it installs", () => {
  it("reuses a row holding only a run status an installed endpoint reports into", async () => {
    const rows = await appendBeneathOneRow({
      ...runItemWithStatusOnly,
      endpoints: {
        runItem_selected: { action: vi.fn(), runStatus: "runStatus" },
      },
    });

    expect(rows).toEqual([[expect.stringMatching(/^r:rit:/), "appended", ""]]);
  });

  it("appends beneath that row when no endpoints are given", async () => {
    const rows = await appendBeneathOneRow(runItemWithStatusOnly);

    expect(rows).toEqual([
      [null, null, "Succeeded"],
      [expect.stringMatching(/^r:rit:/), "appended", ""],
    ]);
  });

  it("appends beneath that row when no given endpoint reports into its column", async () => {
    const rows = await appendBeneathOneRow({
      ...runItemWithStatusOnly,
      endpoints: { runItem_selected: { action: vi.fn(), runStatus: "result" } },
    });

    expect(rows).toHaveLength(2);
  });
});
