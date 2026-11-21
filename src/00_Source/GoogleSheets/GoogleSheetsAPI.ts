import { Val } from "@byronbroughten/utils/val";

import { PartialTableRefusal } from "../RawSource/PartialTableRefusal.ts";
import {
  type DropdownStyleReset,
  type GridFetchOptions,
  type GridFetchRange,
  type LocalWriteOperation,
  type OpaqueRawRequest,
  type RawSource,
  refusesDropdownStyleReset,
  type SheetConditionalFormatSnapshot,
  type SheetEditProtectionSnapshot,
  type SheetPropertiesFetchOptions,
  type SpreadsheetSnapshot,
} from "../RawSource/RawSource.ts";
import { RuleFlushRefusal } from "../RawSource/RuleFlushRefusal.ts";
import { AppsScript } from "./AppsScript.ts";
import { googleConditionalFormatRule } from "./GoogleSheetsAPI/conditionalFormats.ts";
import { googleGrid } from "./GoogleSheetsAPI/gridSnapshots.ts";
import { httpSheetsTransport } from "./GoogleSheetsAPI/httpSheetsTransport.ts";
import { localOperationToGoogleRequests } from "./GoogleSheetsAPI/operationRequests.ts";
import { googleProtectedRange } from "./GoogleSheetsAPI/protectedRanges.ts";

export type GoogleRequest = GoogleAppsScript.Sheets.Schema.Request;

export function googleRawRequest(request: GoogleRequest): OpaqueRawRequest {
  return request as unknown as OpaqueRawRequest;
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
  | "cutPaste"
  | "findReplace"
  | "sortRange"
  | "addConditionalFormatRule"
  | "updateConditionalFormatRule"
  | "deleteConditionalFormatRule"
  | "addProtectedRange"
  | "deleteProtectedRange"
  | "addSheet"
  | "deleteSheet"
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

const partOfTableRefusalText =
  "You cannot insert or delete cells over part of a table.";

const ruleRequestRefusal =
  /requests\[\d+\]\.(add|update)ConditionalFormatRule:/;

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
  refusesDropdownStyleReset: boolean;
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
 * RawSource. HTTP transport lives in GoogleSheetsAPI/ so Node does not install a Sheets global.
 * Translation by subject: GoogleSheetsAPI/ cellData, conditionalFormats,
 * protectedRanges, gridSnapshots, googleColor, operationRequests,
 * httpSheetsTransport.
 * docs/architecture/round-trips.md, queued-writes.md, how-it-runs/node-host.md
 */
export class GoogleSheetsAPI implements RawSource {
  private sheets: SheetsAdvancedTransport;
  private spreadsheetId: string;
  private gridMasks: GridMasks;
  readonly refusesDropdownStyleReset: boolean;
  constructor(props: GoogleSheetsAPIProps) {
    this.sheets = props.sheets;
    this.spreadsheetId = props.spreadsheetId;
    this.gridMasks = gridMasksByHost[props.host];
    this.refusesDropdownStyleReset = props.refusesDropdownStyleReset;
  }
  static init(props: GoogleSheetsAPIProps): GoogleSheetsAPI {
    return new GoogleSheetsAPI(props);
  }
  static forAppsScript(
    dropdownStyleReset: DropdownStyleReset | undefined,
  ): GoogleSheetsAPI {
    const spreadsheetId = AppsScript.boundSpreadsheetId();
    return GoogleSheetsAPI.init({
      sheets: Val.assert(Sheets, "Sheets (enable the Advanced Sheets Service)"),
      spreadsheetId,
      host: "appsScript",
      refusesDropdownStyleReset: refusesDropdownStyleReset(dropdownStyleReset),
    });
  }
  // The Node host's gate judges a dropdown reset against the setting, so nothing is refused here.
  static initHttp(props: GoogleSheetsAPIHttpProps): GoogleSheetsAPI {
    return GoogleSheetsAPI.init({
      sheets: httpSheetsTransport(props),
      spreadsheetId: props.spreadsheetId,
      host: "nodeHost",
      refusesDropdownStyleReset: false,
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
      throw (
        partialTableRefusal(error, requestOperations) ??
        ruleFlushRefusal(error) ??
        error
      );
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

// Observed: docs/sheets-behaviour.md#writing-conditional-formats.
function ruleFlushRefusal(error: unknown): RuleFlushRefusal | undefined {
  if (!(error instanceof Error)) return undefined;
  if (!ruleRequestRefusal.test(error.message)) return undefined;
  return new RuleFlushRefusal(error.message);
}
