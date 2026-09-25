import { GoogleSheetsAPI } from "../00_Source/GoogleSheets/GoogleSheetsAPI";
import { installRawSource } from "../00_Source/RawSource/RawSource";
import { ssConfigGet } from "../01_SpreadsheetSchema/spreadsheetConfigTypes";

type BatchUpdateRequest =
  GoogleAppsScript.Sheets.Schema.BatchUpdateSpreadsheetRequest;
type BatchUpdateResponse =
  GoogleAppsScript.Sheets.Schema.BatchUpdateSpreadsheetResponse;
type GoogleCellData = GoogleAppsScript.Sheets.Schema.CellData;

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
}
export type FakeCell = FakeCellValue | FakeRichCellValue;

export interface FakeSheetProperties {
  sheetId: number;
  title: string;
  /**
   * Row-major grid data, starting at row/column 0 — row indexes here are
   * literal sheet row indexes, so they must line up with
   * `spreadsheetConfig`'s row layout (row 0 is the columnId row, row 4 is
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
   * row; appending a row increments it in place, matching production
   * (`RowRaw.append`).
   */
  table?: {
    endRowIndex: number;
    /**
     * The exclusive bound of the Table's columns, defaulting to the widest
     * row in `rows`. Set it narrower to model the real payload's habit of
     * describing every grid column while the Table covers only some of them.
     */
    endColumnIndex?: number;
    /**
     * Where the Table's range starts, defaulting to the layout every sheet
     * is required to follow (`tableHeaderRowIndexBase0`/`startTableColIndexBase0`).
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
  };
  extraTables?: readonly {
    endRowIndex: number;
    endColumnIndex?: number;
    startRowIndex?: number;
    startColumnIndex?: number;
    name?: string;
  }[]; // Extra Tables besides `table`; the filter hatch still withholds every Table.
  /**
   * The sheet's conditional format rules, in Sheets order. Returned by both
   * `get` and `getByDataFilter`. Adds insert at the requested index and
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
}

export interface FakeSheetsService {
  /** Every request object passed to `Sheets.Spreadsheets.batchUpdate`, in call order. */
  batchUpdateCalls: BatchUpdateRequest[];
  /** Every resource object passed to `Sheets.Spreadsheets.getByDataFilter`, in call order. */
  getByDataFilterCalls: object[];
  /** The params of every `Sheets.Spreadsheets.get` call, in call order. */
  getCalls: { fields?: string }[];
}

/**
 * Builds a `FakeSheetProperties["rows"]` array from a sparse `{ rowIndex:
 * cells }` map, padding the gaps with empty rows so array position lines
 * up with literal sheet row index (row 0 is the columnId row, row 4 is
 * the first data row, per `spreadsheetConfig` — see `rows`' own doc).
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
    data.userEnteredValue = { formulaValue: "=FAKE_FORMULA()" };
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
}: FakeSheetProperties):
  GoogleAppsScript.Sheets.Schema.Sheet["data"] | undefined {
  if (!rows) {
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

function tableStartColumnIndex(
  table: NonNullable<FakeSheetProperties["table"]>,
): number {
  return table.startColumnIndex ?? ssConfigGet("startTableColIndexBase0");
}

function fakeCellHeader(cell: FakeCell | undefined): string | undefined {
  if (cell === null || cell === undefined) return undefined;
  const value = typeof cell === "object" ? cell.value : cell;
  return typeof value === "string" && value !== "" ? value : undefined;
}

function fakeTableColumnProperties(
  sheet: FakeSheetProperties,
  table: NonNullable<FakeSheetProperties["table"]>,
): GoogleAppsScript.Sheets.Schema.TableColumnProperties[] | undefined {
  const {
    columnValidationValues = {},
    columnValidationConditionTypes = {},
    columnTypes = {},
  } = table;
  const range = fakeTableRange(sheet, table);
  const startColumnIndex = range.startColumnIndex ?? 0;
  const endColumnIndex = range.endColumnIndex ?? startColumnIndex;
  if (endColumnIndex <= startColumnIndex) {
    return undefined;
  }
  const headerRow = sheet.rows?.[range.startRowIndex ?? 0] ?? [];
  return Array.from(
    { length: endColumnIndex - startColumnIndex },
    (_, tableRelativeIndex) => {
      const colIndex = startColumnIndex + tableRelativeIndex;
      const colProps: GoogleAppsScript.Sheets.Schema.TableColumnProperties = {};
      if (tableRelativeIndex !== 0) {
        colProps.columnIndex = tableRelativeIndex;
      }
      const columnName = fakeCellHeader(headerRow[colIndex]);
      if (columnName !== undefined) {
        colProps.columnName = columnName;
      }
      const columnType = columnTypes[colIndex];
      if (columnType) {
        colProps.columnType = columnType;
      }
      const values = columnValidationValues[colIndex];
      const conditionType = columnValidationConditionTypes[colIndex];
      if (values || conditionType) {
        colProps.dataValidationRule = {
          condition: {
            ...(conditionType ? { type: conditionType } : {}),
            ...(values
              ? {
                  values: values.map((userEnteredValue) => ({
                    userEnteredValue,
                  })),
                }
              : {}),
          },
        };
      }
      return colProps;
    },
  );
}

function fakeTableRange(
  sheet: FakeSheetProperties,
  table: NonNullable<FakeSheetProperties["table"]>,
): GoogleAppsScript.Sheets.Schema.GridRange {
  return {
    startRowIndex:
      table.startRowIndex ?? ssConfigGet("tableHeaderRowIndexBase0"),
    endRowIndex: table.endRowIndex,
    startColumnIndex: tableStartColumnIndex(table),
    endColumnIndex:
      table.endColumnIndex ??
      Math.max(0, ...(sheet.rows ?? []).map((row) => row.length)),
  };
}

function fakeSheetTables(
  sheet: FakeSheetProperties,
  isFilteredFetch: boolean,
): GoogleAppsScript.Sheets.Schema.Table[] | undefined {
  const { table, extraTables = [] } = sheet;
  if (!table) {
    return undefined;
  }
  if (sheet.isTableHiddenFromFilteredFetch && isFilteredFetch) {
    return undefined;
  }
  return [
    {
      tableId: `fake-table-${sheet.sheetId}`,
      ...(table.name !== undefined ? { name: table.name } : {}),
      range: fakeTableRange(sheet, table),
      columnProperties: fakeTableColumnProperties(sheet, table),
    },
    ...extraTables.map((extraTable, extraIndex) => ({
      tableId: `fake-table-${sheet.sheetId}-extra-${extraIndex}`,
      ...(extraTable.name !== undefined ? { name: extraTable.name } : {}),
      range: fakeTableRange(sheet, extraTable),
    })),
  ];
}

export function stubSheetsService(
  options: FakeSheetsServiceOptions = {},
): FakeSheetsService {
  const sheets = options.sheets ?? [];
  const timeZone =
    options.timeZone === undefined ? fakeTimeZone : options.timeZone;
  const batchUpdateCalls: BatchUpdateRequest[] = [];
  const getByDataFilterCalls: object[] = [];
  const getCalls: { fields?: string }[] = [];
  let nextProtectedRangeId =
    Math.max(
      0,
      ...sheets.flatMap((sheet) =>
        (sheet.protectedRanges ?? []).map(
          (protection) => protection.protectedRangeId ?? 0,
        ),
      ),
    ) + 1;

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
      sheets: sheets.map((s): GoogleAppsScript.Sheets.Schema.Sheet => ({
        properties: { sheetId: s.sheetId, title: s.title },
        data: fakeRowsToGoogleSheetData(s),
        tables: fakeSheetTables(s, isFilteredFetch),
        ...(includeConditionalFormats && s.conditionalFormats?.length
          ? { conditionalFormats: s.conditionalFormats }
          : {}),
        ...(includeProtectedRanges && s.protectedRanges?.length
          ? { protectedRanges: s.protectedRanges }
          : {}),
      })),
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
        const replies = (resource.requests ?? []).map((request) =>
          replaySheetRequest(sheets, request, () => nextProtectedRangeId++),
        );
        return { replies };
      },
    },
  };

  installRawSource(GoogleSheetsAPI.init(service, fakeSpreadsheetId));

  return { batchUpdateCalls, getByDataFilterCalls, getCalls };
}

function replaySheetRequest(
  sheets: FakeSheetProperties[],
  request: GoogleAppsScript.Sheets.Schema.Request,
  nextProtectedRangeId: () => number,
): GoogleAppsScript.Sheets.Schema.Response {
  replayConditionalFormatRequest(sheets, request);
  replayUpdateTableRequest(sheets, request);
  return replayProtectedRangeRequest(sheets, request, nextProtectedRangeId);
}

// Like the live API: the sent list replaces the Table's, and every column needs a name.
function replayUpdateTableRequest(
  sheets: FakeSheetProperties[],
  request: GoogleAppsScript.Sheets.Schema.Request,
): void {
  const update = request.updateTable;
  const table = update?.table;
  if (table?.tableId === undefined || update?.fields !== "columnProperties") {
    return;
  }
  const sheet = sheets.find(
    (candidate) => `fake-table-${candidate.sheetId}` === table.tableId,
  );
  if (sheet?.table === undefined) return;
  const tableState = sheet.table;
  const startColumnIndex = tableStartColumnIndex(tableState);
  const headerRowIndex =
    tableState.startRowIndex ?? ssConfigGet("tableHeaderRowIndexBase0");
  const columnProperties = table.columnProperties ?? [];
  columnProperties.forEach((column) => {
    if (column.columnName === undefined || column.columnName === "") {
      throw new Error(
        `Invalid requests[0].updateTable: column ${column.columnIndex ?? 0} has no columnName.`,
      );
    }
  });
  tableState.columnTypes = {};
  tableState.columnValidationValues = {};
  tableState.columnValidationConditionTypes = {};
  const {
    columnTypes,
    columnValidationValues,
    columnValidationConditionTypes,
  } = tableState;
  const rows = [...(sheet.rows ?? [])];
  const headerRow = [...(rows[headerRowIndex] ?? [])];
  columnProperties.forEach((column) => {
    const colIndex = startColumnIndex + (column.columnIndex ?? 0);
    headerRow[colIndex] = column.columnName ?? null;
    if (column.columnType !== undefined) {
      columnTypes[colIndex] = column.columnType;
    }
    const condition = column.dataValidationRule?.condition;
    if (condition?.type !== undefined) {
      columnValidationConditionTypes[colIndex] = condition.type;
    }
    const values = (condition?.values ?? []).flatMap((value) =>
      value.userEnteredValue === undefined ? [] : [value.userEnteredValue],
    );
    if (values.length > 0) columnValidationValues[colIndex] = values;
  });
  rows[headerRowIndex] = Array.from(headerRow, (cell) => cell ?? null);
  sheet.rows = rows;
}

function replayProtectedRangeRequest(
  sheets: FakeSheetProperties[],
  request: GoogleAppsScript.Sheets.Schema.Request,
  nextProtectedRangeId: () => number,
): GoogleAppsScript.Sheets.Schema.Response {
  const add = request.addProtectedRange?.protectedRange;
  if (add !== undefined) {
    const sheetId = add.range?.sheetId;
    const sheet = sheets.find((candidate) => candidate.sheetId === sheetId);
    if (sheet === undefined) return {};
    const stored = {
      ...add,
      protectedRangeId: nextProtectedRangeId(),
    };
    sheet.protectedRanges ??= [];
    sheet.protectedRanges.push(stored);
    return { addProtectedRange: { protectedRange: stored } };
  }
  const removeId = request.deleteProtectedRange?.protectedRangeId;
  if (removeId === undefined) return {};
  sheets.forEach((sheet) => {
    if (sheet.protectedRanges === undefined) return;
    sheet.protectedRanges = sheet.protectedRanges.filter(
      (protection) => protection.protectedRangeId !== removeId,
    );
  });
  return {};
}

function replayConditionalFormatRequest(
  sheets: FakeSheetProperties[],
  request: GoogleAppsScript.Sheets.Schema.Request,
): void {
  const add = request.addConditionalFormatRule;
  if (add?.rule !== undefined) {
    const sheetId = add.rule.ranges?.[0]?.sheetId;
    const sheet = sheets.find((candidate) => candidate.sheetId === sheetId);
    if (sheet === undefined) return;
    sheet.conditionalFormats ??= [];
    sheet.conditionalFormats.splice(add.index ?? 0, 0, add.rule);
    return;
  }
  const remove = request.deleteConditionalFormatRule;
  if (remove?.index === undefined) return;
  const sheet = sheets.find(
    (candidate) => candidate.sheetId === remove.sheetId,
  );
  if (sheet?.conditionalFormats === undefined) return;
  sheet.conditionalFormats.splice(remove.index, 1);
}
