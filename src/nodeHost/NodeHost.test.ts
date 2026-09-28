import { describe, expect, it, vi } from "vitest";

import type { SheetsHttpRequest } from "../00_Source/GoogleSheets/GoogleSheetsAPI";
import { getColumnTraitByName } from "../01_SpreadsheetSchema/columnConfigsTypes";
import { installedConfigs } from "../01_SpreadsheetSchema/configRegister";
import { getSheetTraitByName } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import { sheetLayout } from "../01_SpreadsheetSchema/sheetLayout";
import { SpreadsheetRaw } from "../02_SpreadsheetRaw/SpreadsheetRaw";
import type { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import type { Endpoints } from "../06_API/Endpoints";
import {
  buildGridRows,
  type FakeCell,
  type FakeCellValue,
} from "../testSupport/fakeSheetsService";
import { NodeHost } from "./NodeHost";

const spreadsheetId = "spreadsheet-under-test";
const gadgetsGid = 111;

const gadgetsPayload = {
  sheets: [
    {
      properties: { sheetId: gadgetsGid, title: "Gadgets" },
      tables: [
        {
          tableId: "fake-table",
          range: {
            startRowIndex: sheetLayout.tableHeaderRowIndex,
            endRowIndex: 11,
            startColumnIndex: sheetLayout.startTableColIndex,
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
  raw.sheet(gadgetsGid).row(5).cell(2).updateValue("Processing...");
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
  sheetGid: getSheetTraitByName("runItem", "sheetGid"),
  title: "Run item",
  columnIds: [
    getColumnTraitByName("runItem", "id", "columnId"),
    getColumnTraitByName("runItem", "result", "columnId"),
    getColumnTraitByName("runItem", "runStatus", "columnId"),
  ],
  topRow: [null, null, "Succeeded"],
  append: (ss: SpreadsheetNamed) => {
    const sheet = ss.sheet("runItem");
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

  it("counts the framework's own feedback columns on the config sheets with no endpoints given", async () => {
    const rows = await appendBeneathOneRow({
      sheetGid: getSheetTraitByName("spreadsheetConfig", "sheetGid"),
      title: "Spreadsheet Config",
      columnIds: [
        getColumnTraitByName("spreadsheetConfig", "tableMenuSpace", "columnId"),
        getColumnTraitByName(
          "spreadsheetConfig",
          "fillRowIdsRunStatus",
          "columnId",
        ),
      ],
      topRow: [null, "Succeeded"],
      append: (ss) => {
        const sheet = ss.sheet("spreadsheetConfig");
        sheet.row(topDataRowIndex).prepFetchFull();
        ss.fetchAllPrepped();
        sheet.appendRowWithVals({ tableMenuSpace: "space" });
      },
    });

    expect(rows).toEqual([["space", ""]]);
  });
});
