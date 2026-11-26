import type {
  SheetsHttpRequest,
  SheetsHttpTransport,
} from "../00_Source/GoogleSheets/GoogleSheetsAPI.ts";
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
  params: [string, string][];
  body: string | null;
}

// Keyed by method and URL suffix; anything else goes unanswered.
const routes: Record<string, (call: FakeCall) => unknown> = {
  "GET "({ fake, spreadsheetId, params }) {
    return fake.sheets.Spreadsheets.get(spreadsheetId, fieldsArg(params));
  },
  "POST :getByDataFilter"({ fake, spreadsheetId, params, body }) {
    const resource = parsedBody(body, ["dataFilters", "includeGridData"]);
    validateGridFilters(resource);
    return fake.sheets.Spreadsheets.getByDataFilter(
      resource,
      spreadsheetId,
      fieldsArg(params),
    );
  },
  "POST :batchUpdate"({ fake, spreadsheetId, params, body }) {
    validateKeys(
      params.map(([key]) => key),
      [],
      "query parameter",
    );
    const resource: BatchUpdateRequest = parsedBody(body, ["requests"]);
    return fake.sheets.Spreadsheets.batchUpdate(resource, spreadsheetId);
  },
  "GET /values:batchGet"({ fake, spreadsheetId, params }) {
    return valuesBatchGet.reply({ grid: fake.grid, spreadsheetId, params });
  },
};

// Answers the Node host's requests from the spreadsheet a `stubSheetsService` fake holds.
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
  return route({ fake, spreadsheetId, params: queryParams(query), body });
}

function queryParams(query: string | undefined): [string, string][] {
  if (query === undefined) return [];
  return query.split("&").map((pair) => {
    const [key = "", ...value] = pair.split("=");
    return [decoded(key), decoded(value.join("="))];
  });
}

function decoded(text: string): string {
  return decodeURIComponent(text.replaceAll("+", " "));
}

function fieldsArg(params: [string, string][]): { fields?: string } {
  validateKeys(
    params.map(([key]) => key),
    ["fields"],
    "query parameter",
  );
  return Object.fromEntries(params);
}

function parsedBody(body: string | null, keys: string[]): object {
  if (body === null) throw unreplayed("a POST with no body");
  const resource = JSON.parse(body) as object;
  validateKeys(Object.keys(resource), keys, "body key");
  return resource;
}

function validateKeys(
  keys: string[],
  modelled: string[],
  kind: "query parameter" | "body key",
): void {
  const other = keys.find((key) => !modelled.includes(key));
  if (other !== undefined) throw unreplayed(`the ${kind} "${other}"`);
}

// The stub drops any other filter without a word, so its Tables would silently go missing.
function validateGridFilters(resource: object): void {
  const { dataFilters = [] } = resource as { dataFilters?: object[] };
  if (dataFilters.some((filter) => !("gridRange" in filter))) {
    throw unreplayed("a dataFilter without a gridRange");
  }
}
