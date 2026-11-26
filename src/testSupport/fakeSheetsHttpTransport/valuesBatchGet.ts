import type { FakeCell, FakeCellValue } from "../fakeSheetsService.ts";
import { a1Notation } from "../fakeSheetsService/a1Notation.ts";
import { fakeCells } from "../fakeSheetsService/fakeCells.ts";
import type { BoundedRange } from "../fakeSheetsService/fakeGrid.ts";
import type {
  FakeGridView,
  FakeSheetView,
} from "../fakeSheetsService/gridView.ts";
import { unreplayed } from "../fakeSheetsService/unreplayed.ts";

type ValueRange = GoogleAppsScript.Sheets.Schema.ValueRange;
type MajorDimension = "ROWS" | "COLUMNS";
type ValueRender = "UNFORMATTED_VALUE" | "FORMULA";

// Groups: quoted title, plain title, start column, start row, end column, end row.
const a1Range =
  /^(?:'((?:[^']|'')+)'|([A-Za-z_]\w*))!([A-Z]{1,3})([1-9]\d*)?(?::([A-Z]{1,3})([1-9]\d*)?)?$/;

interface ValuesBatchGetProps {
  grid: FakeGridView;
  spreadsheetId: string;
  params: [string, string][];
}

interface ValuesOptions {
  ranges: string[];
  majorDimension: MajorDimension;
  valueRender: ValueRender;
}

// `spreadsheets.values.batchGet`, in the request and response shape of Google's Sheets API.
export const valuesBatchGet = {
  reply({ grid, spreadsheetId, params }: ValuesBatchGetProps): {
    spreadsheetId: string;
    valueRanges: ValueRange[];
  } {
    const options = valuesOptions(params);
    return {
      spreadsheetId,
      valueRanges: options.ranges.map((range) =>
        valueRange(grid, range, options),
      ),
    };
  },
};

function valuesOptions(params: [string, string][]): ValuesOptions {
  const ranges: string[] = [];
  let majorDimension: MajorDimension = "ROWS";
  // Google's default, which the fake can't answer, as it formats nothing.
  let valueRender = "FORMATTED_VALUE";
  for (const [key, value] of params) {
    if (key === "ranges") {
      ranges.push(value);
    } else if (key === "valueRenderOption") {
      valueRender = value;
    } else if (key === "majorDimension" && isMajorDimension(value)) {
      majorDimension = value;
    } else if (key !== "dateTimeRenderOption" || value !== "SERIAL_NUMBER") {
      throw unreplayed(`the query parameter "${key}" as "${value}"`);
    }
  }
  if (ranges.length === 0) {
    throw new Error("values.batchGet needs at least one range.");
  }
  if (valueRender !== "UNFORMATTED_VALUE" && valueRender !== "FORMULA") {
    throw unreplayed(`valueRenderOption ${valueRender}`);
  }
  return { ranges, majorDimension, valueRender };
}

function isMajorDimension(value: string): value is MajorDimension {
  return value === "ROWS" || value === "COLUMNS";
}

function valueRange(
  grid: FakeGridView,
  range: string,
  { majorDimension, valueRender }: ValuesOptions,
): ValueRange {
  const { sheet, bounds } = parsedRange(grid, range);
  const rows = sheet
    .rows(bounds)
    .map((row) => row.map((cell) => rendered(cell, valueRender)));
  let lines = rows;
  if (majorDimension === "COLUMNS") lines = transposed(rows);
  const values = withoutTrailing(
    lines.map((line) => withoutTrailing(line, (value) => value === "")),
    (line) => line.length === 0,
  );
  return {
    range: echoedRange(sheet.title, bounds),
    majorDimension,
    // Like Google: a range holding nothing comes back with no values.
    ...(values.length > 0 ? { values } : {}),
  };
}

// A whole column or an open end is bounded by the grid; whole rows and a bare sheet name aren't modelled.
function parsedRange(
  grid: FakeGridView,
  range: string,
): { sheet: FakeSheetView; bounds: BoundedRange } {
  const match = a1Range.exec(range);
  if (match === null) throw unreplayed(`the range "${range}"`);
  const [, quoted, plain = "", startColumn = "", startRow] = match;
  const [, , , , , endColumn, endRow] = match;
  // A lone column with no row, such as Title!A, is no A1 range.
  if (endColumn === undefined && startRow === undefined) {
    throw unreplayed(`the range "${range}"`);
  }
  const sheet = grid.sheetByTitle(quoted?.replaceAll("''", "'") ?? plain);
  let lastRow = endRow;
  if (endColumn === undefined) lastRow = startRow;
  const bounds = {
    startRowIndex: Number(startRow ?? 1) - 1,
    endRowIndex: Number(lastRow ?? sheet.rowCount),
    startColumnIndex: a1Notation.columnIndex(startColumn),
    endColumnIndex: a1Notation.columnIndex(endColumn ?? startColumn) + 1,
  };
  // Live may clip such a range instead, which is unmeasured.
  if (
    bounds.startRowIndex >= bounds.endRowIndex ||
    bounds.startColumnIndex >= bounds.endColumnIndex ||
    bounds.endRowIndex > sheet.rowCount ||
    bounds.endColumnIndex > sheet.columnCount
  ) {
    throw unreplayed(`the range "${range}", reversed or past the grid`);
  }
  return { sheet, bounds };
}

function rendered(
  cell: FakeCell,
  valueRender: ValueRender,
): NonNullable<FakeCellValue> {
  if (valueRender === "FORMULA") {
    const formula = fakeCells.formula(cell);
    if (formula !== undefined) return formula;
  }
  return fakeCells.value(cell) ?? "";
}

function transposed<IT>(rows: IT[][]): IT[][] {
  return (rows[0] ?? []).map((_, colIndex) =>
    rows.flatMap((row) => row.slice(colIndex, colIndex + 1)),
  );
}

function withoutTrailing<IT>(
  items: IT[],
  isEmpty: (item: IT) => boolean,
): IT[] {
  return items.slice(0, items.findLastIndex((item) => !isEmpty(item)) + 1);
}

function echoedRange(title: string, bounds: BoundedRange): string {
  const cells = a1Notation.range(bounds);
  const [first, last] = cells.split(":");
  const echoed = first === last ? first : cells;
  return `${a1Notation.sheetName(title)}!${echoed}`;
}
