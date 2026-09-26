import {
  GoogleSheetsAPI,
  type ModeledRequestVerb,
} from "../00_Source/GoogleSheets/GoogleSheetsAPI";
import { installRawSource } from "../00_Source/RawSource/RawSource";
import { Obj } from "../utils/Obj";
import { cellReplays } from "./fakeSheetsService/cellReplays";
import { dimensionReplays } from "./fakeSheetsService/dimensionReplays";
import {
  type FakeSheetState,
  type FakeSpreadsheet,
  fakeSpreadsheet,
} from "./fakeSheetsService/fakeSpreadsheet";
import { fakeTables } from "./fakeSheetsService/fakeTables";
import { type FakeGridView, gridView } from "./fakeSheetsService/gridView";
import { ruleReplays } from "./fakeSheetsService/ruleReplays";
import { sheetReplays } from "./fakeSheetsService/sheetReplays";
import { tableReplays } from "./fakeSheetsService/tableReplays";

type BatchUpdateRequest =
  GoogleAppsScript.Sheets.Schema.BatchUpdateSpreadsheetRequest;
type BatchUpdateResponse =
  GoogleAppsScript.Sheets.Schema.BatchUpdateSpreadsheetResponse;
type GoogleCellData = GoogleAppsScript.Sheets.Schema.CellData;
type Request = GoogleAppsScript.Sheets.Schema.Request;
type Response = GoogleAppsScript.Sheets.Schema.Response;

export const fakeSpreadsheetId = "fake-spreadsheet";
export const fakeTimeZone = "America/Chicago";

/** A single cell's value, in the same terms `CellRaw` reads/writes them.
 * `null` (or a short row) represents an empty cell. */
export type FakeCellValue = string | number | boolean | null;

/**
 * A cell's value plus the live facts `ColumnConfigOperator
 * .fetchAndUpdateColumnConfig` reads off it (`CellRaw.isFormula`/
 * `numberFormatType`) — use this richer form instead of a bare
 * `FakeCellValue` wherever a test needs to mark a cell as a live formula or
 * give it a number format (e.g. `DATE`) distinguishable from a plain number.
 */
export interface FakeRichCellValue {
  value: FakeCellValue;
  isFormula?: boolean;
  numberFormatType?: string;
  dataValidationConditionType?: string;
  backgroundColor?: GoogleAppsScript.Sheets.Schema.Color;
}
export type FakeCell = FakeCellValue | FakeRichCellValue;

export interface FakeTable {
  /** Omit for `fake-table-<gid>` (`-extra-<i>` for an extra Table). */
  tableId?: string;
  endRowIndex: number;
  /**
   * The exclusive bound of the Table's columns, defaulting to the widest
   * row in `rows`. Set it narrower to model the real payload's habit of
   * describing every grid column while the Table covers only some of them.
   */
  endColumnIndex?: number;
  /**
   * Where the Table's range starts, defaulting to the layout every sheet
   * is required to follow (`sheetLayout.tableHeaderRowIndex`/`startTableColIndex`).
   * Override either one only to build a deliberately misplaced Table, which
   * `SpreadsheetRaw`'s post-fetch placement check refuses.
   */
  startRowIndex?: number;
  startColumnIndex?: number;
  /**
   * The Table's name in Sheets. Omit to leave it unnamed in the payload,
   * which the adapter stores as "".
   */
  name?: string;
  /**
   * A column's live data-validation condition values (e.g.
   * `["=valueConfig[Transaction Description]"]`), keyed by absolute
   * column index — read by `ColumnConfigOperator`'s valueName detection
   * (`ColumnMetaRaw.valueValidationStrings`). Omit for a table with no
   * validated columns.
   */
  columnValidationValues?: Record<number, string[]>;
  /**
   * A column's live data-validation condition type (e.g. `"BOOLEAN"`
   * from Insert > Checkbox), keyed by absolute column index. A BOOLEAN
   * rule typically has no `values`, so this is independent of
   * `columnValidationValues`.
   */
  columnValidationConditionTypes?: Record<number, string>;
  /**
   * A column's declared Sheets column type (e.g. `"CURRENCY"`, `"DATE"`,
   * `"BOOLEAN"`), keyed by absolute column index — read by
   * `ColumnMetaRaw.activeColumnType`, which `ColumnConfigOperator`'s
   * valueName derivation consults before falling back to the top-row
   * sample. Omit a column here to leave it untyped (Automatic), which is
   * what the real API reports for a column whose type was never set.
   * GET/replay convert this key to Google's table-relative `columnIndex`.
   */
  columnTypes?: Record<number, string>;
}

export interface FakeSheetProperties {
  sheetId: number;
  title: string;
  /**
   * Row-major grid data, starting at row/column 0 — row indexes here are
   * literal sheet row indexes, so they must line up with
   * `sheetLayout`'s row layout (row 0 is the columnId row, row 4 is
   * the first data row, etc.) for anything above 02_SpreadsheetRaw to
   * resolve columns/values correctly. Omit for a sheet whose cell content
   * doesn't matter to the test (sheet-properties-only fixtures still work
   * as before).
   */
  rows?: readonly (readonly FakeCell[])[];
  /**
   * The sheet's Table range. Required for any test that reads/appends
   * *data* rows on this sheet (`SheetRaw`'s `rowIndexesActive`/
   * `appendDataRow` etc. read `activeTable`, which throws if no table was
   * ever integrated) — not needed for sheets only read via a uniform row
   * (e.g. a business sheet's header row). `endRowIndex` is the exclusive
   * bound of existing data rows and must be strictly past the first data
   * row; a replayed Table append grows it, as the live API does.
   */
  table?: FakeTable;
  extraTables?: readonly FakeTable[]; // Extra Tables besides `table`; the filter hatch still withholds every Table.
  /**
   * The sheet's conditional format rules, in Sheets order. Returned by
   * `get` only, as live. Adds insert at the requested index and
   * deletes remove-and-renumber, so a later read sees the list a flush left.
   */
  conditionalFormats?: GoogleAppsScript.Sheets.Schema.ConditionalFormatRule[];
  /**
   * The sheet's protected ranges. Omitted from `getByDataFilter` and from
   * a `get` that does not ask for them; an empty list is omitted, like Google.
   */
  protectedRanges?: GoogleAppsScript.Sheets.Schema.ProtectedRange[];
  /**
   * Row indices that come back as a grid-data block describing every column
   * but carrying no `rowData` at all — the real API's shape, measured
   * against the live spreadsheet, for a row inside the grid whose every
   * cell lacks a value, a formula and an explicit number format. Contrast
   * a row present in `rows` with empty cells, which IS reported and so
   * marks those cells active with an empty value. Use this to reproduce
   * bugs where code assumes a row it explicitly fetched came back with
   * cells in the response.
   */
  rowsWithNoGridData?: readonly number[];
  /**
   * Row indices that come back with no grid-data block at all, splitting
   * the grid into separate blocks — what the API really does for a row
   * past the populated grid, as opposed to a blank row inside it.
   */
  rowsWithNoGridBlock?: readonly number[];
  /**
   * Makes this sheet's Table come back from `Spreadsheets.get` but not from
   * `getByDataFilter`, reproducing the real API's rule that a sheet's
   * `tables` metadata is returned only for a filter whose range overlaps the
   * table — the blind spot a Table that moved down or right falls into. A
   * deliberate escape hatch, not filter awareness: teaching the fake real
   * range arithmetic would put a second, subtly wrong model of the Sheets
   * API into test support.
   */
  isTableHiddenFromFilteredFetch?: boolean;
}

export interface FakeSheetsServiceOptions {
  sheets?: FakeSheetProperties[];
  /** The spreadsheet's `properties.timeZone`; `null` leaves it out of every response. */
  timeZone?: string | null;
  /** Records and counts batch updates without replaying them, as the Node host's dry run sends none. */
  isDryRun?: boolean;
}

export interface FakeSheetsService {
  /** Every request object passed to `Sheets.Spreadsheets.batchUpdate`, in call order. */
  batchUpdateCalls: BatchUpdateRequest[];
  /** How many batch updates were sent, for a test whose name states a round-trip cost. */
  batchUpdateCount(): number;
  /** The fake spreadsheet as every replayed batch update left it. */
  grid: FakeGridView;
  /** Every resource object passed to `Sheets.Spreadsheets.getByDataFilter`, in call order. */
  getByDataFilterCalls: object[];
  /** The params of every `Sheets.Spreadsheets.get` call, in call order. */
  getCalls: { fields?: string }[];
}

/**
 * Builds a `FakeSheetProperties["rows"]` array from a sparse `{ rowIndex:
 * cells }` map, padding the gaps with empty rows so array position lines
 * up with literal sheet row index (row 0 is the columnId row, row 4 is
 * the first data row, per `sheetLayout` — see `rows`' own doc).
 */
export function buildGridRows(
  rowsByIndex: Record<number, readonly FakeCell[]>,
): FakeCell[][] {
  const maxRowIndex = Math.max(0, ...Object.keys(rowsByIndex).map(Number));
  return Array.from({ length: maxRowIndex + 1 }, (_, rowIndex) => [
    ...(rowsByIndex[rowIndex] ?? []),
  ]);
}

function fakeCellToGoogleCellData(cell: FakeCell): GoogleCellData {
  if (cell === null) {
    return {};
  }
  const rich: FakeRichCellValue =
    typeof cell === "object" ? cell : { value: cell };
  const data: GoogleCellData = {
    effectiveValue: fakeValueToExtendedValue(rich.value),
  };
  if (rich.isFormula) {
    data.userEnteredValue = { formulaValue: formulaText(rich.value) };
  }
  if (rich.numberFormatType) {
    data.effectiveFormat = {
      numberFormat: { type: rich.numberFormatType },
    };
  }
  if (rich.dataValidationConditionType) {
    data.dataValidation = {
      condition: { type: rich.dataValidationConditionType },
    };
  }
  return data;
}

// A replayed formula keeps its text; a fixture's formula cell gives only its computed value.
function formulaText(value: FakeCellValue): string {
  return typeof value === "string" && value.startsWith("=")
    ? value
    : "=FAKE_FORMULA()";
}

function fakeValueToExtendedValue(
  value: FakeCellValue,
): GoogleAppsScript.Sheets.Schema.ExtendedValue | undefined {
  if (value === null) {
    return undefined;
  }
  if (typeof value === "string") {
    return { stringValue: value };
  }
  if (typeof value === "number") {
    return { numberValue: value };
  }
  return { boolValue: value };
}

function fakeRowsToGoogleSheetData({
  rows,
  rowsWithNoGridData = [],
  rowsWithNoGridBlock = [],
}: FakeSheetState): GoogleAppsScript.Sheets.Schema.Sheet["data"] | undefined {
  if (rows.length === 0) {
    return undefined;
  }
  const columnCount = Math.max(0, ...rows.map((row) => row.length));
  const noGridDataRows = new Set(rowsWithNoGridData);
  const noGridBlockRows = new Set(rowsWithNoGridBlock);

  // Real rowData blocks run contiguously from a real startRow, so either
  // kind of absent row ends the block it was in.
  const blocks: NonNullable<GoogleAppsScript.Sheets.Schema.Sheet["data"]> = [];
  let currentBlockRows: GoogleAppsScript.Sheets.Schema.RowData[] = [];
  let currentBlockStart: number | null = null;
  function columnMetadata(): GoogleAppsScript.Sheets.Schema.DimensionProperties[] {
    return Array.from({ length: columnCount }, () => ({}));
  }
  function flushCurrentBlock(): void {
    if (currentBlockStart !== null) {
      blocks.push({
        startColumn: 0,
        startRow: currentBlockStart,
        columnMetadata: columnMetadata(),
        rowData: currentBlockRows,
      });
    }
    currentBlockRows = [];
    currentBlockStart = null;
  }
  rows.forEach((row, rowIndex) => {
    if (noGridBlockRows.has(rowIndex)) {
      flushCurrentBlock();
      return;
    }
    if (noGridDataRows.has(rowIndex)) {
      flushCurrentBlock();
      // Every grid column described, no rowData — the measured real shape.
      blocks.push({
        startColumn: 0,
        startRow: rowIndex,
        columnMetadata: columnMetadata(),
      });
      return;
    }
    currentBlockStart ??= rowIndex;
    currentBlockRows.push({
      values: Array.from(
        { length: columnCount },
        (_, colIndex): GoogleCellData =>
          fakeCellToGoogleCellData(row[colIndex] ?? null),
      ),
    });
  });
  flushCurrentBlock();
  return blocks;
}

export function stubSheetsService(
  options: FakeSheetsServiceOptions = {},
): FakeSheetsService {
  const spreadsheet = fakeSpreadsheet.init(options.sheets ?? []);
  const timeZone =
    options.timeZone === undefined ? fakeTimeZone : options.timeZone;
  const batchUpdateCalls: BatchUpdateRequest[] = [];
  const getByDataFilterCalls: object[] = [];
  const getCalls: { fields?: string }[] = [];

  function sheetsResponse(
    isFilteredFetch: boolean,
    fields?: string,
  ): GoogleAppsScript.Sheets.Schema.Spreadsheet {
    // Like Google: getByDataFilter drops rules, and an empty list is omitted.
    const includeConditionalFormats =
      !isFilteredFetch &&
      (fields === undefined || fields.includes("conditionalFormats"));
    const includeProtectedRanges =
      !isFilteredFetch &&
      (fields === undefined || fields.includes("protectedRanges"));
    const includeTimeZone =
      timeZone !== null &&
      (fields === undefined || fields.includes("timeZone"));
    return {
      ...(includeTimeZone ? { properties: { timeZone } } : {}),
      sheets: spreadsheet.sheets.map(
        (s): GoogleAppsScript.Sheets.Schema.Sheet => ({
          properties: { sheetId: s.sheetId, title: s.title },
          data: fakeRowsToGoogleSheetData(s),
          tables: fakeTables.googleTables(s, isFilteredFetch),
          ...(includeConditionalFormats && s.conditionalFormats?.length
            ? { conditionalFormats: s.conditionalFormats }
            : {}),
          ...(includeProtectedRanges && s.protectedRanges?.length
            ? { protectedRanges: s.protectedRanges }
            : {}),
        }),
      ),
    };
  }

  const service = {
    Spreadsheets: {
      get: (_spreadsheetId: string, params?: { fields?: string }) => {
        getCalls.push(params ?? {});
        return sheetsResponse(false, params?.fields);
      },
      getByDataFilter: (
        resource: object,
        _spreadsheetId: string,
        params?: { fields?: string },
      ) => {
        getByDataFilterCalls.push(resource);
        return sheetsResponse(true, params?.fields);
      },
      batchUpdate: (
        resource: BatchUpdateRequest,
        _spreadsheetId: string,
      ): BatchUpdateResponse => {
        batchUpdateCalls.push(resource);
        if (options.isDryRun) return {};
        return replayBatch(spreadsheet, resource.requests ?? []);
      },
    },
  };

  installRawSource(GoogleSheetsAPI.init(service, fakeSpreadsheetId));

  return {
    batchUpdateCalls,
    batchUpdateCount() {
      return batchUpdateCalls.length;
    },
    grid: gridView.build(spreadsheet),
    getByDataFilterCalls,
    getCalls,
  };
}

// Every kind the adapter sends, plus the raw-request kinds a chore reaches for.
const requestReplays = {
  ...cellReplays,
  ...dimensionReplays,
  ...sheetReplays,
  ...tableReplays,
  ...ruleReplays,
} satisfies Record<ModeledRequestVerb, unknown>;

type ReplayedKind = keyof typeof requestReplays;
type RequestReplay = (spreadsheet: FakeSpreadsheet, body: unknown) => Response;

// Like the live API, a batch applies in order and all or nothing.
function replayBatch(
  spreadsheet: FakeSpreadsheet,
  requests: Request[],
): BatchUpdateResponse {
  const before = JSON.stringify(spreadsheet);
  try {
    return {
      replies: requests.map((request) => replayRequest(spreadsheet, request)),
    };
  } catch (error) {
    Object.assign(spreadsheet, JSON.parse(before));
    throw error;
  }
}

function replayRequest(
  spreadsheet: FakeSpreadsheet,
  request: Request,
): Response {
  const kinds = Obj.keys(request).filter((kind) => request[kind] !== undefined);
  const [kind] = kinds;
  if (kinds.length !== 1 || kind === undefined) {
    throw new Error(
      `A request names exactly one kind, not ${kinds.length}: ${kinds.join(", ")}.`,
    );
  }
  if (!isReplayedKind(kind)) {
    throw new Error(`The fake Sheets service does not replay ${kind}.`);
  }
  // Each replay takes its own kind's body, which the lookup can't express.
  const replay = requestReplays[kind] as RequestReplay;
  return replay(spreadsheet, request[kind]);
}

function isReplayedKind(kind: string): kind is ReplayedKind {
  return kind in requestReplays;
}
