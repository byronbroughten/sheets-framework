import type {
  SheetsHttpRequest,
  SheetsHttpTransport,
} from "../00_Source/GoogleSheets/GoogleSheetsAPI.ts";
import { FakeQueryParams } from "./fakeSheetsHttpTransport/FakeQueryParams.ts";
import { validateKeys } from "./fakeSheetsHttpTransport/validateKeys.ts";
import { valuesBatchGet } from "./fakeSheetsHttpTransport/valuesBatchGet.ts";
import type { FakeSheetsService } from "./fakeSheetsService.ts";
import { unreplayed } from "./fakeSheetsService/unreplayed.ts";

type BatchUpdateRequest =
  GoogleAppsScript.Sheets.Schema.BatchUpdateSpreadsheetRequest;

// Parsed by hand: `src/` has no URL class. Groups: spreadsheet ID, suffix, query.
const spreadsheetUrl =
  /^https:\/\/sheets\.googleapis\.com\/v4\/spreadsheets\/([^/:?]+)([^?]*)(?:\?(.*))?$/;

interface FakeCall {
  fake: FakeSheetsService;
  spreadsheetId: string;
  params: FakeQueryParams;
  body: string | null;
}

// Keyed by method and URL suffix; anything else goes unanswered.
const routes: Record<string, (call: FakeCall) => unknown> = {
  "GET "({ fake, spreadsheetId, params }) {
    return fake.sheets.Spreadsheets.get(spreadsheetId, params.fieldsArg);
  },
  "POST :getByDataFilter"({ fake, spreadsheetId, params, body }) {
    const resource = parsedBody(body, ["dataFilters", "includeGridData"]);
    validateGridData(resource);
    validateGridFilters(resource);
    return fake.sheets.Spreadsheets.getByDataFilter(
      resource,
      spreadsheetId,
      params.fieldsArg,
    );
  },
  "POST :batchUpdate"({ fake, spreadsheetId, params, body }) {
    params.validateKeys([]);
    const resource: BatchUpdateRequest = parsedBody(body, ["requests"]);
    return fake.sheets.Spreadsheets.batchUpdate(resource, spreadsheetId);
  },
  "GET /values:batchGet"({ fake, spreadsheetId, params }) {
    return valuesBatchGet.reply({ grid: fake.grid, spreadsheetId, params });
  },
};

export function fakeSheetsHttpTransport(
  fake: FakeSheetsService,
): SheetsHttpTransport {
  return (request) => {
    try {
      // As on the wire: an undefined drops out, and nothing is shared with the fake's state.
      return JSON.parse(JSON.stringify(reply(fake, request))) as unknown;
    } catch (error) {
      const detail = error instanceof Error ? error.message : String(error);
      throw new Error(
        `The fake Sheets transport can't answer ${request.method} ${request.url}: ${detail}`,
        { cause: error },
      );
    }
  };
}

function reply(
  fake: FakeSheetsService,
  { method, url, body }: SheetsHttpRequest,
): unknown {
  const [, spreadsheetId, suffix = "", query] = spreadsheetUrl.exec(url) ?? [];
  const route = routes[`${method} ${suffix}`];
  if (spreadsheetId === undefined || route === undefined) {
    throw unreplayed("that method and URL");
  }
  const params = FakeQueryParams.init(query);
  return route({ fake, spreadsheetId, params, body });
}

function parsedBody(body: string | null, modelled: string[]): object {
  if (body === null) throw unreplayed("a POST with no body");
  const resource = JSON.parse(body) as object;
  validateKeys({ keys: Object.keys(resource), modelled, kind: "body key" });
  return resource;
}

// The stub sends grid data whatever this says.
function validateGridData(resource: object): void {
  const { includeGridData } = resource as { includeGridData?: boolean };
  if (includeGridData !== true) {
    throw unreplayed("a getByDataFilter without includeGridData");
  }
}

// The stub drops any other filter without a word, so its Tables would silently go missing.
function validateGridFilters(resource: object): void {
  const { dataFilters = [] } = resource as { dataFilters?: object[] };
  if (dataFilters.some((filter) => !("gridRange" in filter))) {
    throw unreplayed("a dataFilter without a gridRange");
  }
}
