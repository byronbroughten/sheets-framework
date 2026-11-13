import { Val } from "@byronbroughten/utils/val";

import { PartialTableRefusal } from "../RawSource/PartialTableRefusal";
import type {
  GridFetchOptions,
  GridFetchRange,
  LocalWriteOperation,
  OpaqueRawRequest,
  RawSource,
  SheetConditionalFormatSnapshot,
  SheetEditProtectionSnapshot,
  SheetPropertiesFetchOptions,
  SpreadsheetSnapshot,
  TableColumnPropertiesAdd,
  TableColumnPropertiesUpdate,
  UpdateTabPropertiesOperation,
} from "../RawSource/RawSource";
import { AppsScript } from "./AppsScript";
import { cellDataRequests } from "./GoogleSheetsAPI/cellData";
import { googleConditionalFormatRule } from "./GoogleSheetsAPI/conditionalFormats";
import { googleColor } from "./GoogleSheetsAPI/googleColor";
import { googleGrid } from "./GoogleSheetsAPI/gridSnapshots";
import { googleProtectedRange } from "./GoogleSheetsAPI/protectedRanges";

export type GoogleRequest = GoogleAppsScript.Sheets.Schema.Request;

export function googleRawRequest(request: GoogleRequest): OpaqueRawRequest {
  return request as unknown as OpaqueRawRequest;
}

function unwrapRawRequest(request: OpaqueRawRequest): GoogleRequest {
  return request as unknown as GoogleRequest;
}

// Naming a verb here obliges UpdateRequestSummary to give it a line format and FlushClassifier to judge it.
export type ModeledRequestVerb =
  | "appendCells"
  | "appendDimension"
  | "insertDimension"
  | "deleteDimension"
  | "insertRange"
  | "deleteRange"
  | "copyPaste"
  | "findReplace"
  | "sortRange"
  | "addConditionalFormatRule"
  | "deleteConditionalFormatRule"
  | "addProtectedRange"
  | "deleteProtectedRange"
  | "addSheet"
  | "addTable"
  | "deleteTable"
  | "updateSheetProperties"
  | "updateTable"
  | "updateCells"
  | "setDataValidation"
  | "repeatCell"
  | "pasteData";

export type ModeledRequest = {
  [RV in ModeledRequestVerb]: Required<Pick<GoogleRequest, RV>>;
}[ModeledRequestVerb];

export type GoogleBatchUpdateResponse =
  GoogleAppsScript.Sheets.Schema.BatchUpdateSpreadsheetResponse;

type GoogleSpreadsheet = GoogleAppsScript.Sheets.Schema.Spreadsheet;
type GoogleSheet = GoogleAppsScript.Sheets.Schema.Sheet;
type GoogleGridRange = GoogleAppsScript.Sheets.Schema.GridRange;
type BatchUpdateRequest =
  GoogleAppsScript.Sheets.Schema.BatchUpdateSpreadsheetRequest;
type GetByDataFilterRequest =
  GoogleAppsScript.Sheets.Schema.GetSpreadsheetByDataFilterRequest;

interface FieldsArg {
  fields?: string;
}

const sheetsApiBase = "https://sheets.googleapis.com/v4/spreadsheets";

const partOfTableRefusalText =
  "You cannot insert or delete cells over part of a table.";

const timeZoneMask = "properties(timeZone)";
const sheetPropertiesMask =
  "properties(sheetId,title,hidden,tabColorStyle,gridProperties(rowCount,columnCount,frozenRowCount,frozenColumnCount))";

const sheetPropertiesAndTablesMask =
  `${sheetPropertiesMask},` +
  "tables(tableId,name,range,columnProperties(columnIndex,columnType))";

// The time zone, grid size and column types ride every standing fetch, so none costs its own get.
const fieldMasks = {
  timeZone: timeZoneMask,
  sheetProperties: `${timeZoneMask},sheets(${sheetPropertiesAndTablesMask})`,
  sheetPropertiesWithConditionalFormats: `${timeZoneMask},sheets(${sheetPropertiesAndTablesMask},conditionalFormats)`,
  conditionalFormats: "sheets(properties(sheetId),conditionalFormats)",
  protectedRanges: "sheets(properties(sheetId),protectedRanges)",
  gridWithProgrammaticFacts:
    `${timeZoneMask},` +
    "sheets(" +
    `${sheetPropertiesMask},` +
    "tables(tableId,name,range,columnProperties(columnIndex,columnName,columnType,dataValidationRule(condition(type,values(userEnteredValue)))))," +
    "data(startColumn,startRow,columnMetadata,rowData(values(effectiveValue,userEnteredValue,effectiveFormat(numberFormat(type)),dataValidation(condition(type)))))" +
    ")",
  gridEffectiveValuesOnly: plainGridMask("effectiveValue"),
  // A non-formula cell comes back as userEnteredValue {}; observed: docs/sheets-behaviour.md#reading-formula-text-beside-effective-values.
  gridWithFormulas: plainGridMask(
    "effectiveValue,userEnteredValue(formulaValue)",
  ),
} as const;

function plainGridMask(cellFields: string): string {
  return (
    `${timeZoneMask},` +
    "sheets(" +
    `${sheetPropertiesMask},` +
    "tables(tableId,name,range,columnProperties(columnIndex,columnType))," +
    `data(startColumn,startRow,columnMetadata,rowData(values(${cellFields})))` +
    ")"
  );
}

interface GridMasks {
  withProgrammaticFacts: string;
  withoutProgrammaticFacts: string;
}

// The Node host's gate judges entered data per cell, so its default read carries each cell's formula.
const gridMasksByHost: Record<"appsScript" | "nodeHost", GridMasks> = {
  appsScript: {
    withProgrammaticFacts: fieldMasks.gridWithProgrammaticFacts,
    withoutProgrammaticFacts: fieldMasks.gridEffectiveValuesOnly,
  },
  nodeHost: {
    withProgrammaticFacts: fieldMasks.gridWithProgrammaticFacts,
    withoutProgrammaticFacts: fieldMasks.gridWithFormulas,
  },
};

export interface GoogleSheetsAPIProps {
  sheets: SheetsAdvancedTransport;
  spreadsheetId: string;
  host: keyof typeof gridMasksByHost;
}

export interface SheetsHttpRequest {
  method: "GET" | "POST";
  url: string;
  body: string | null;
}

export type SheetsHttpTransport = (request: SheetsHttpRequest) => unknown;

export interface SheetsHttpRead {
  // Google trims trailing empty rows and cells from a reply, so what a read covers comes from the ranges it asked for.
  gridRanges: GoogleGridRange[];
  spreadsheet: GoogleSpreadsheet;
}

export interface GoogleSheetsAPIHttpProps {
  spreadsheetId: string;
  transport: SheetsHttpTransport;
  reportRead: (read: SheetsHttpRead) => void;
  // Sends the flush by calling send, or withholds it by not calling it.
  gateFlush: (
    requests: GoogleRequest[],
    send: () => GoogleBatchUpdateResponse,
  ) => GoogleBatchUpdateResponse;
}

export interface SheetsAdvancedTransport {
  Spreadsheets: {
    get: (spreadsheetId: string, optionalArgs?: FieldsArg) => GoogleSpreadsheet;
    getByDataFilter: (
      resource: GetByDataFilterRequest,
      spreadsheetId: string,
      optionalArgs?: FieldsArg,
    ) => GoogleSpreadsheet;
    batchUpdate: (
      resource: BatchUpdateRequest,
      spreadsheetId: string,
    ) => GoogleBatchUpdateResponse;
  };
}

/**
 * Google adapter for RawSource: with its GoogleSheetsAPI/ subfolder, the only
 * code that may construct or consume Google Sheets schema types or call
 * Spreadsheets.get / getByDataFilter / batchUpdate. SpreadsheetRaw holds a
 * RawSource. HTTP transport lives here so Node does not install a Sheets global.
 * Translation by subject: GoogleSheetsAPI/ cellData, conditionalFormats,
 * protectedRanges, gridSnapshots, googleColor.
 * docs/architecture/round-trips.md, queued-writes.md, how-it-runs/node-host.md
 */
export class GoogleSheetsAPI implements RawSource {
  private sheets: SheetsAdvancedTransport;
  private spreadsheetId: string;
  private gridMasks: GridMasks;
  constructor({ sheets, spreadsheetId, host }: GoogleSheetsAPIProps) {
    this.sheets = sheets;
    this.spreadsheetId = spreadsheetId;
    this.gridMasks = gridMasksByHost[host];
  }
  static init(props: GoogleSheetsAPIProps): GoogleSheetsAPI {
    return new GoogleSheetsAPI(props);
  }
  static forAppsScript(): GoogleSheetsAPI {
    const spreadsheetId = AppsScript.boundSpreadsheetId();
    return GoogleSheetsAPI.init({
      sheets: Val.assert(Sheets, "Sheets (enable the Advanced Sheets Service)"),
      spreadsheetId,
      host: "appsScript",
    });
  }
  static initHttp(props: GoogleSheetsAPIHttpProps): GoogleSheetsAPI {
    return GoogleSheetsAPI.init({
      sheets: httpSheetsTransport(props),
      spreadsheetId: props.spreadsheetId,
      host: "nodeHost",
    });
  }
  fetchSheetProperties(
    options?: SheetPropertiesFetchOptions,
  ): SpreadsheetSnapshot {
    if (options?.includeConditionalFormats) {
      return googleGrid.toSpreadsheetSnapshotWithRules(
        this._getSpreadsheet(fieldMasks.sheetPropertiesWithConditionalFormats),
        googleConditionalFormatRule.toConditionalFormatRules,
      );
    }
    return googleGrid.toSpreadsheetSnapshot(
      this._getSpreadsheet(fieldMasks.sheetProperties),
    );
  }
  fetchTimeZone(): string | null {
    const spreadsheet = this._getSpreadsheet(fieldMasks.timeZone);
    return spreadsheet.properties?.timeZone ?? null;
  }
  fetchGrid(
    gridRanges: GridFetchRange[],
    options: GridFetchOptions,
  ): SpreadsheetSnapshot {
    return googleGrid.toSpreadsheetSnapshot(
      this.sheets.Spreadsheets.getByDataFilter(
        {
          dataFilters: gridRanges.map((gr) => ({ gridRange: gr })),
          includeGridData: true,
        },
        this.spreadsheetId,
        {
          fields: this._gridFields(options),
        },
      ),
    );
  }
  // getByDataFilter returned no conditionalFormats, so rules take a plain get; observed: docs/sheets-behaviour.md#reading-conditional-formats.
  fetchConditionalFormatRules(): SheetConditionalFormatSnapshot[] {
    return this._fetchSheets(fieldMasks.conditionalFormats).map((sheet) => ({
      sheetGid: sheetGidOf(sheet),
      rules: googleConditionalFormatRule.toConditionalFormatRules(sheet),
    }));
  }
  // The code assumes Google drops an empty protectedRanges list the same way; observed: docs/sheets-behaviour.md#reading-protections-back.
  fetchEditProtections(): SheetEditProtectionSnapshot[] {
    return this._fetchSheets(fieldMasks.protectedRanges).map((sheet) => ({
      sheetGid: sheetGidOf(sheet),
      protections: (sheet.protectedRanges ?? []).map(
        googleProtectedRange.toEditProtection,
      ),
    }));
  }
  flush(operations: LocalWriteOperation[]): void {
    const sentRequests = operations.flatMap((operation) =>
      localOperationToGoogleRequests(operation).map((request) => ({
        request,
        operation,
      })),
    );
    const requests = sentRequests.map(({ request }) => request);
    const requestOperations = sentRequests.map(({ operation }) => operation);
    if (requests.length === 0) return;
    try {
      const response = this.sheets.Spreadsheets.batchUpdate(
        { requests },
        this.spreadsheetId,
      );
      googleProtectedRange.validateAddReplies(response, requests);
    } catch (error) {
      throw partialTableRefusal(error, requestOperations) ?? error;
    }
  }
  // The Sheets service has no toast, so this needs the Apps Script host; only an endpoint run toasts.
  toast(message: string): void {
    AppsScript.toast(message);
  }
  private _gridFields(options: GridFetchOptions): string {
    return options.includeProgrammaticFacts
      ? this.gridMasks.withProgrammaticFacts
      : this.gridMasks.withoutProgrammaticFacts;
  }
  private _fetchSheets(fields: string): GoogleSheet[] {
    return Val.assert(
      this._getSpreadsheet(fields).sheets,
      "spreadsheet.sheets",
    );
  }
  private _getSpreadsheet(fields: string): GoogleSpreadsheet {
    return this.sheets.Spreadsheets.get(this.spreadsheetId, { fields });
  }
}

function sheetGidOf(sheet: GoogleSheet): number {
  return Val.assert(sheet.properties?.sheetId, "sheetId");
}

// Google prefixes its refusal with the request, "Invalid requests[3].insertRange: …"; observed: docs/sheets-behaviour.md#inserting-and-deleting-cells-beside-and-inside-a-table.
function partialTableRefusal(
  error: unknown,
  requestOperations: LocalWriteOperation[],
): PartialTableRefusal | undefined {
  if (!(error instanceof Error)) return undefined;
  if (!error.message.includes(partOfTableRefusalText)) return undefined;
  const requestIndex = /requests\[(\d+)\]/.exec(error.message)?.[1];
  if (requestIndex === undefined) return undefined;
  const operation = requestOperations[Number(requestIndex)];
  if (operation === undefined) return undefined;
  return new PartialTableRefusal(error.message, operation);
}

function httpSheetsTransport(
  props: GoogleSheetsAPIHttpProps,
): SheetsAdvancedTransport {
  // The one place the wire is trusted, as Apps Script's own declaration trusts it.
  function send<RS>(request: SheetsHttpRequest): RS {
    return props.transport(request) as RS;
  }
  function url(spreadsheetId: string, suffix: string, fields?: string): string {
    const query = fields ? `?fields=${encodeURIComponent(fields)}` : "";
    return `${sheetsApiBase}/${spreadsheetId}${suffix}${query}`;
  }
  return {
    Spreadsheets: {
      get: (spreadsheetId, optionalArgs) => {
        const spreadsheet = send<GoogleSpreadsheet>({
          method: "GET",
          url: url(spreadsheetId, "", optionalArgs?.fields),
          body: null,
        });
        props.reportRead({ gridRanges: [], spreadsheet });
        return spreadsheet;
      },
      getByDataFilter: (resource, spreadsheetId, optionalArgs) => {
        const spreadsheet = send<GoogleSpreadsheet>({
          method: "POST",
          url: url(spreadsheetId, ":getByDataFilter", optionalArgs?.fields),
          body: JSON.stringify(resource),
        });
        props.reportRead({
          gridRanges: (resource.dataFilters ?? []).flatMap(({ gridRange }) =>
            gridRange === undefined ? [] : [gridRange],
          ),
          spreadsheet,
        });
        return spreadsheet;
      },
      batchUpdate: (resource, spreadsheetId) =>
        props.gateFlush(resource.requests ?? [], () =>
          send<GoogleBatchUpdateResponse>({
            method: "POST",
            url: url(spreadsheetId, ":batchUpdate"),
            body: JSON.stringify(resource),
          }),
        ),
    },
  };
}

function localOperationToGoogleRequests(
  operation: LocalWriteOperation,
): GoogleRequest[] {
  if (operation.kind === "raw") return [unwrapRawRequest(operation.request)];
  return modeledOperationToGoogleRequests(operation);
}

function modeledOperationToGoogleRequests(
  operation: Exclude<LocalWriteOperation, { kind: "raw" }>,
): ModeledRequest[] {
  switch (operation.kind) {
    case "appendDimension":
      return [
        {
          appendDimension: {
            sheetId: operation.sheetId,
            dimension: operation.dimension,
            length: operation.addedCount,
          },
        },
      ];
    case "insertRange":
      return [
        {
          insertRange: {
            range: operation.range,
            shiftDimension: operation.shiftDimension,
          },
        },
      ];
    case "updateTableRange":
      return [
        {
          updateTable: {
            table: { tableId: operation.tableId, range: operation.range },
            fields: "range",
          },
        },
      ];
    case "deleteRange":
      return [
        {
          deleteRange: {
            range: operation.range,
            shiftDimension: operation.shiftDimension,
          },
        },
      ];
    case "copyPaste":
      return [
        {
          copyPaste: {
            source: operation.source,
            destination: operation.destination,
            pasteType: operation.pasteType,
            pasteOrientation: "NORMAL",
          },
        },
      ];
    case "fillColumn":
      return cellDataRequests.fillColumn(operation);
    case "fillCell":
      return cellDataRequests.fillCell(operation);
    case "findReplace":
      return [
        {
          findReplace: { ...operation.terms, ...operation.scope },
        },
      ];
    case "deleteTableRows":
      return [
        {
          deleteRange: { range: operation.range, shiftDimension: "ROWS" },
        },
      ];
    case "sortTable":
      return [
        {
          sortRange: {
            range: operation.range,
            sortSpecs: [
              {
                dimensionIndex: operation.colIdxToSortBy,
                sortOrder: operation.sortOrder,
              },
            ],
          },
        },
      ];
    case "addConditionalFormatRule":
      return [
        {
          addConditionalFormatRule: {
            index: operation.index,
            rule: googleConditionalFormatRule.fromModelable(operation.rule),
          },
        },
      ];
    case "deleteConditionalFormatRule":
      return [
        {
          deleteConditionalFormatRule: {
            sheetId: operation.sheetId,
            index: operation.index,
          },
        },
      ];
    case "addProtectedRange":
      return [
        {
          addProtectedRange: {
            protectedRange: googleProtectedRange.fromContent(
              operation.protection,
            ),
          },
        },
      ];
    case "deleteProtectedRange":
      return [
        {
          deleteProtectedRange: {
            protectedRangeId: operation.protectedRangeId,
          },
        },
      ];
    case "addSheet":
      return [
        {
          addSheet: {
            properties: {
              sheetId: operation.sheetId,
              title: operation.title,
              gridProperties: {
                rowCount: operation.rowCount,
                columnCount: operation.columnCount,
              },
            },
          },
        },
      ];
    // Google writes addTable's own column names three columns right of the Table; observed: docs/sheets-behaviour.md#creating-and-updating-tables.
    case "addTable":
      return [
        {
          addTable: {
            table: {
              tableId: operation.tableId,
              name: operation.name,
              range: operation.range,
            },
          },
        },
        columnPropertiesRequest(operation.tableId, operation.columnProperties),
      ];
    // No rows, so the head-row cells blank and their checkboxes lose validation; observed: docs/sheets-behaviour.md#creating-and-updating-tables.
    case "deleteTable":
      return [
        { deleteTable: { tableId: operation.tableId } },
        {
          updateCells: {
            range: operation.headRows,
            fields: "userEnteredValue,dataValidation",
          },
        },
      ];
    case "renameSheet":
      return [
        {
          updateSheetProperties: {
            properties: {
              sheetId: operation.sheetId,
              title: operation.title,
            },
            fields: "title",
          },
        },
      ];
    case "updateTabProperties":
      return [tabPropertiesRequest(operation)];
    case "renameTable":
      return [
        {
          updateTable: {
            table: {
              tableId: operation.tableId,
              name: operation.name,
            },
            fields: "name",
          },
        },
      ];
    case "setTableColumnProperties":
      return [
        columnPropertiesRequest(operation.tableId, operation.columnProperties),
      ];
    case "addCheckboxValidation":
      return [
        {
          setDataValidation: {
            range: operation.range,
            rule: { condition: { type: "BOOLEAN" } },
          },
        },
      ];
    default: {
      const exhaustive: never = operation;
      throw new Error(
        `Unknown local write operation: ${JSON.stringify(exhaustive)}`,
      );
    }
  }
}

type TabProperty = keyof UpdateTabPropertiesOperation["properties"];

// The field each tab property sets; the mask names only what the operation carries, so an absent property is left alone.
const tabPropertyMasks: [TabProperty, string][] = [
  ["color", "tabColorStyle"],
  ["hidden", "hidden"],
  ["frozenRows", "gridProperties.frozenRowCount"],
  ["frozenColumns", "gridProperties.frozenColumnCount"],
  ["index", "index"],
];

function tabPropertiesRequest({
  sheetId,
  properties,
}: UpdateTabPropertiesOperation): ModeledRequest {
  const { color, hidden, frozenRows, frozenColumns, index } = properties;
  const fields = tabPropertyMasks.flatMap(([property, mask]) =>
    properties[property] === undefined ? [] : [mask],
  );
  return {
    updateSheetProperties: {
      properties: {
        sheetId,
        tabColorStyle:
          color === undefined ? undefined : googleColor.fromTabColor(color),
        hidden,
        gridProperties: {
          frozenRowCount: frozenRows,
          frozenColumnCount: frozenColumns,
        },
        index,
      },
      fields: fields.join(","),
    },
  };
}

function columnPropertiesRequest(
  tableId: string,
  columnProperties: TableColumnPropertiesAdd[] | TableColumnPropertiesUpdate[],
): ModeledRequest {
  return {
    updateTable: {
      table: { tableId, columnProperties },
      fields: "columnProperties",
    },
  };
}
