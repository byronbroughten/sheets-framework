import { describe, expect, it, vi } from "vitest";

import type { SheetsHttpRequest } from "../00_Source/GoogleSheets/GoogleSheetsAPI";
import { installedConfigs } from "../01_SpreadsheetSchema/configRegister";
import { ssConfigGet } from "../01_SpreadsheetSchema/spreadsheetConfigTypes";
import { SpreadsheetRaw } from "../02_SpreadsheetRaw/SpreadsheetRaw";
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
            startRowIndex: ssConfigGet("tableHeaderRowIndexBase0"),
            endRowIndex: 11,
            startColumnIndex: ssConfigGet("startTableColIndexBase0"),
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
