import type {
  GoogleRequest,
  ModeledRequestVerb,
} from "../00_Source/GoogleSheets/GoogleSheetsAPI";
import { sheetConfigsByGid } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import { Obj } from "../utils/Obj";

type GoogleUpdateRequest = GoogleRequest;
type GoogleGridRange = GoogleAppsScript.Sheets.Schema.GridRange;
type UserEnteredValue = NonNullable<
  NonNullable<
    GoogleAppsScript.Sheets.Schema.UpdateCellsRequest["rows"]
  >[number]["values"]
>[number]["userEnteredValue"];
type RowData = GoogleAppsScript.Sheets.Schema.RowData;
type DimensionRange = GoogleAppsScript.Sheets.Schema.DimensionRange;

const layoutLimits = {
  maxSampleValues: 4,
  maxRawChars: 120,
  subjectWidth: 24,
  countWidth: 14,
} as const;
const findReplaceFlags = [
  "matchCase",
  "matchEntireCell",
  "searchByRegex",
  "includeFormulas",
] as const;

type RequestVerb = keyof GoogleUpdateRequest;

export interface UpdateRequestSummaryProps {
  requests: GoogleUpdateRequest[];
}

// One line per request, so the preview is read rather than waved through.
export class UpdateRequestSummary {
  private requests: GoogleUpdateRequest[];
  constructor({ requests }: UpdateRequestSummaryProps) {
    this.requests = requests;
  }
  static init(requests: GoogleUpdateRequest[]): UpdateRequestSummary {
    return new UpdateRequestSummary({ requests });
  }
  get isEmpty(): boolean {
    return this.requests.length === 0;
  }
  get count(): number {
    return this.requests.length;
  }
  get json(): string {
    return JSON.stringify({ requests: this.requests }, null, 2);
  }
  get lines(): string[] {
    return this.requests.map(line);
  }
}

function line(request: GoogleUpdateRequest): string {
  const verb = requestVerb(request);
  if (verb === undefined) return "(empty request)";
  return [verb.padEnd(26), body(verb, request)].join(" ");
}

function requestVerb(request: GoogleUpdateRequest): RequestVerb | undefined {
  return Obj.keys(request)[0];
}

function body(verb: RequestVerb, request: GoogleUpdateRequest): string {
  if (Obj.isKey(requestBody, verb)) return modeledBody(verb, request);
  return rawBody(request, verb);
}

function modeledBody<RV extends ModeledRequestVerb>(
  verb: RV,
  request: GoogleUpdateRequest,
): string {
  return requestBody[verb](request[verb]);
}

const requestBody: {
  [RV in ModeledRequestVerb]: (inner: GoogleUpdateRequest[RV]) => string;
} = {
  updateCells(
    updateCells: GoogleAppsScript.Sheets.Schema.UpdateCellsRequest | undefined,
  ): string {
    return columns(
      label.range(updateCells?.range),
      `${cellCount(updateCells?.rows)} cell(s)`,
      label.values(updateCells?.rows),
      label.fields(updateCells?.fields),
    );
  },
  repeatCell(
    repeatCell: GoogleAppsScript.Sheets.Schema.RepeatCellRequest | undefined,
  ): string {
    const range = repeatCell?.range;
    return columns(
      label.range(range),
      `${rangeCellCount(range)} cell(s)`,
      label.values([{ values: [repeatCell?.cell ?? {}] }]),
      label.fields(repeatCell?.fields),
    );
  },
  setDataValidation(
    setDataValidation:
      GoogleAppsScript.Sheets.Schema.SetDataValidationRequest | undefined,
  ): string {
    const range = setDataValidation?.range;
    return columns(
      label.range(range),
      `${rangeCellCount(range)} cell(s)`,
      setDataValidation?.rule?.condition?.type ?? "",
      "",
    );
  },
  appendCells(
    appendCells: GoogleAppsScript.Sheets.Schema.AppendCellsRequest | undefined,
  ): string {
    return columns(
      label.sheet(appendCells?.sheetId),
      `${appendCells?.rows?.length ?? 0} row(s)`,
      label.values(appendCells?.rows),
      label.fields(appendCells?.fields),
    );
  },
  insertDimension(
    insert: GoogleAppsScript.Sheets.Schema.InsertDimensionRequest | undefined,
  ): string {
    return dimensionBody(insert?.range, "insert");
  },
  deleteDimension(
    remove: GoogleAppsScript.Sheets.Schema.DeleteDimensionRequest | undefined,
  ): string {
    return dimensionBody(remove?.range, "delete");
  },
  sortRange(
    sortRange: GoogleAppsScript.Sheets.Schema.SortRangeRequest | undefined,
  ): string {
    const spec = sortRange?.sortSpecs?.[0];
    const column = colLetters(spec?.dimensionIndex ?? 0);
    const order = (spec?.sortOrder ?? "ASCENDING").toLowerCase();
    return columns(
      label.range(sortRange?.range),
      `by column ${column} ${order}`,
      "",
      "",
    );
  },
  // The count lives in a response a dry run never gets, so the line can't claim one.
  findReplace(
    findReplace: GoogleAppsScript.Sheets.Schema.FindReplaceRequest | undefined,
  ): string {
    return columns(
      label.findReplaceScope(findReplace),
      "matching cells",
      `"${findReplace?.find ?? ""}" → "${findReplace?.replacement ?? ""}"`,
      label.findReplaceFlags(findReplace),
    );
  },
  pasteData(
    pasteData: GoogleAppsScript.Sheets.Schema.PasteDataRequest | undefined,
  ): string {
    const coordinate = pasteData?.coordinate;
    const startRowIndex = coordinate?.rowIndex ?? 0;
    const startColumnIndex = coordinate?.columnIndex ?? 0;
    const rowCount = Math.max(rfc4180RecordCount(pasteData?.data ?? ""), 1);
    const range = coordinate && {
      sheetId: coordinate.sheetId,
      startRowIndex,
      endRowIndex: startRowIndex + rowCount,
      startColumnIndex,
      endColumnIndex: startColumnIndex + 1,
    };
    return columns(
      label.range(range),
      `${rowCount} row(s)`,
      firstRfc4180Field(pasteData?.data ?? "").replaceAll("\n", " "),
      pasteData?.type ?? "",
    );
  },
  addConditionalFormatRule(
    add:
      | GoogleAppsScript.Sheets.Schema.AddConditionalFormatRuleRequest
      | undefined,
  ): string {
    const range = add?.rule?.ranges?.[0];
    const conditionType = add?.rule?.booleanRule?.condition?.type ?? "";
    return columns(label.range(range), "prepend", conditionType, "");
  },
  deleteConditionalFormatRule(
    remove:
      | GoogleAppsScript.Sheets.Schema.DeleteConditionalFormatRuleRequest
      | undefined,
  ): string {
    return columns(
      label.sheet(remove?.sheetId),
      `index ${remove?.index ?? ""}`,
      "",
      "",
    );
  },
  addProtectedRange(
    add: GoogleAppsScript.Sheets.Schema.AddProtectedRangeRequest | undefined,
  ): string {
    const protection = add?.protectedRange;
    const kind = protection?.warningOnly === true ? "warning" : "lock";
    return columns(
      label.range(protection?.range),
      kind,
      protection?.description ?? "",
      "",
    );
  },
  deleteProtectedRange(
    remove:
      GoogleAppsScript.Sheets.Schema.DeleteProtectedRangeRequest | undefined,
  ): string {
    return columns(
      "(no sheet)",
      `id ${remove?.protectedRangeId ?? ""}`,
      "",
      "",
    );
  },
  // The tab isn't fetched yet, so the GID prints as a number rather than a config name.
  addSheet(
    add: GoogleAppsScript.Sheets.Schema.AddSheetRequest | undefined,
  ): string {
    const properties = add?.properties;
    const grid = properties?.gridProperties;
    return columns(
      gidLabel(properties?.sheetId),
      "add tab",
      `${properties?.title ?? ""} (${grid?.rowCount ?? 0} rows × ${grid?.columnCount ?? 0} cols)`,
      "",
    );
  },
  addTable(
    add: GoogleAppsScript.Sheets.Schema.AddTableRequest | undefined,
  ): string {
    const table = add?.table;
    const columnProperties = table?.columnProperties;
    return columns(
      label.range(table?.range, gidLabel(table?.range?.sheetId)),
      `add ${columnProperties?.length ?? 0} cols`,
      `${table?.name ?? ""} (${label.tableColumns(columnProperties)})`,
      "",
    );
  },
  updateTable(
    update: GoogleAppsScript.Sheets.Schema.UpdateTableRequest | undefined,
  ): string {
    const table = update?.table;
    if (update?.fields === "name") {
      return columns(
        table?.tableId ?? "(no table)",
        "name",
        table?.name ?? "",
        label.fields(update.fields),
      );
    }
    return columns(
      table?.tableId ?? "(no table)",
      `${table?.columnProperties?.length ?? 0} cols`,
      label.tableColumns(table?.columnProperties),
      label.fields(update?.fields),
    );
  },
  updateSheetProperties(
    update:
      GoogleAppsScript.Sheets.Schema.UpdateSheetPropertiesRequest | undefined,
  ): string {
    const properties = update?.properties;
    return columns(
      label.sheet(properties?.sheetId),
      "title",
      properties?.title ?? "",
      label.fields(update?.fields),
    );
  },
};

// The opening's own line format: no type layer to read it through.
// The opening's own line format: no type layer to read it through.
function rawBody(request: GoogleUpdateRequest, verb: RequestVerb): string {
  const inner = request[verb];
  const json = JSON.stringify(inner ?? {});
  return columns(
    label.sheet(sheetIdOf(inner)),
    "",
    json.length > layoutLimits.maxRawChars
      ? `${json.slice(0, layoutLimits.maxRawChars)}…`
      : json,
    "",
  );
}

function dimensionBody(
  range: DimensionRange | undefined,
  verb: "insert" | "delete",
): string {
  const dimension = (range?.dimension ?? "ROWS").toLowerCase();
  const startIndex = range?.startIndex ?? 0;
  const endIndex = range?.endIndex ?? startIndex;
  const span = label.dimensionSpan(dimension, startIndex, endIndex);
  return columns(
    `${label.sheet(range?.sheetId)}!${span}`,
    `${verb} ${endIndex - startIndex} ${dimension}`,
    "",
    "",
  );
}

function columns(
  subject: string,
  count: string,
  values: string,
  fields: string,
): string {
  return [
    subject.padEnd(layoutLimits.subjectWidth),
    count.padEnd(layoutLimits.countWidth),
    values,
    fields,
  ]
    .join(" ")
    .trimEnd();
}

const label = {
  // An absent bound really is open-ended, so it renders open rather than as one cell.
  range(
    range: GoogleGridRange | undefined,
    sheetLabelOverride?: string,
  ): string {
    if (!range) return "(no range)";
    const sheetLabel = sheetLabelOverride ?? label.sheet(range.sheetId);
    const hasBound =
      range.startRowIndex !== undefined ||
      range.endRowIndex !== undefined ||
      range.startColumnIndex !== undefined ||
      range.endColumnIndex !== undefined;
    if (!hasBound) {
      return `${sheetLabel}!sheet`;
    }
    const { startRowIndex = 0, endRowIndex, endColumnIndex } = range;
    const startCell = `${colLetters(range.startColumnIndex ?? 0)}${startRowIndex + 1}`;
    const endColumn =
      endColumnIndex === undefined ? "" : colLetters(endColumnIndex - 1);
    return `${sheetLabel}!${startCell}:${endColumn}${endRowIndex ?? ""}`;
  },
  tableColumns(
    columnProperties:
      GoogleAppsScript.Sheets.Schema.TableColumnProperties[] | undefined,
  ): string {
    return [...(columnProperties ?? [])]
      .sort((a, b) => (a.columnIndex ?? 0) - (b.columnIndex ?? 0))
      .map(
        (column) =>
          `${column.columnName ?? `col ${column.columnIndex ?? 0}`}: ${column.columnType ?? "—"}`,
      )
      .join(", ");
  },
  values(rows: RowData[] | undefined): string {
    const values = (rows ?? []).flatMap((row) =>
      (row.values ?? []).map((cell) => valueLabel(cell.userEnteredValue)),
    );
    if (values.length === 0) return "(no values)";
    const shown = values.slice(0, layoutLimits.maxSampleValues).join(", ");
    return values.length > layoutLimits.maxSampleValues ? `${shown}, …` : shown;
  },
  fields(fields: string | undefined): string {
    return fields ? `[${fields}]` : "";
  },
  sheet(sheetGid: number | undefined): string {
    if (sheetGid === undefined) return "(no sheet)";
    return sheetConfigsByGid().get(sheetGid)?.sheetName ?? `gid ${sheetGid}`;
  },
  dimensionSpan(
    dimension: string,
    startIndex: number,
    endIndex: number,
  ): string {
    if (dimension === "columns") {
      return `${colLetters(startIndex)}:${colLetters(endIndex - 1)}`;
    }
    return `${startIndex + 1}:${endIndex}`;
  },
  findReplaceScope(
    findReplace: GoogleAppsScript.Sheets.Schema.FindReplaceRequest | undefined,
  ): string {
    if (findReplace?.allSheets) return "every sheet";
    if (findReplace?.sheetId !== undefined) {
      return `${label.sheet(findReplace.sheetId)}!all`;
    }
    return label.range(findReplace?.range);
  },
  findReplaceFlags(
    findReplace: GoogleAppsScript.Sheets.Schema.FindReplaceRequest | undefined,
  ): string {
    const flags = findReplaceFlags.filter((flag) => findReplace?.[flag]);
    if (flags.length === 0) return "(no flags)";
    return flags.join(", ");
  },
};

function gidLabel(sheetGid: number | undefined): string {
  return sheetGid === undefined ? "(no sheet)" : `gid ${sheetGid}`;
}

function cellCount(rows: RowData[] | undefined): number {
  return (rows ?? []).reduce(
    (count, row) => count + (row.values?.length ?? 0),
    0,
  );
}

function rangeCellCount(range: GoogleGridRange | undefined): number {
  const {
    startRowIndex = 0,
    endRowIndex = 0,
    startColumnIndex = 0,
  } = range ?? {};
  const endColumnIndex = range?.endColumnIndex ?? startColumnIndex + 1;
  return (endRowIndex - startRowIndex) * (endColumnIndex - startColumnIndex);
}

function colLetters(colIndex: number): string {
  let remaining = colIndex + 1;
  let letters = "";
  while (remaining > 0) {
    const rest = (remaining - 1) % 26;
    letters = String.fromCharCode(65 + rest) + letters;
    remaining = Math.floor((remaining - 1) / 26);
  }
  return letters;
}

function rfc4180RecordCount(data: string): number {
  if (data === "") return 0;
  let count = 0;
  let i = 0;
  while (i < data.length) {
    count += 1;
    if (data[i] === '"') {
      i += 1;
      while (i < data.length) {
        if (data[i] === '"' && data[i + 1] === '"') {
          i += 2;
          continue;
        }
        if (data[i] === '"') {
          i += 1;
          break;
        }
        i += 1;
      }
    } else {
      while (i < data.length && data[i] !== "\n") i += 1;
    }
    if (data[i] === "\n") i += 1;
  }
  return count;
}

function firstRfc4180Field(data: string): string {
  if (!data.startsWith('"')) return data.split(/\t|\n/, 1)[0] ?? "";
  let field = "";
  for (let i = 1; i < data.length; i++) {
    if (data[i] !== '"') {
      field += data[i];
      continue;
    }
    if (data[i + 1] === '"') {
      field += '"';
      i += 1;
      continue;
    }
    return field;
  }
  return field;
}

// An unmodelled request may still name a sheet, and usually does.
function sheetIdOf(inner: object | undefined): number | undefined {
  const sheetId = (inner as { sheetId?: unknown } | undefined)?.sheetId;
  return typeof sheetId === "number" ? sheetId : undefined;
}

function valueLabel(userEnteredValue: UserEnteredValue): string {
  if (!userEnteredValue) return "(no value)";
  const { stringValue, numberValue, boolValue, formulaValue } =
    userEnteredValue;
  if (stringValue !== undefined) return `"${stringValue}"`;
  if (numberValue !== undefined) return String(numberValue);
  if (boolValue !== undefined) return String(boolValue);
  if (formulaValue !== undefined) return formulaValue;
  return "(no value)";
}
