import type { CellValue } from "../CellValues/cellValues";
import type {
  ConditionalFormatRule,
  ModelableConditionalFormatRule,
} from "./ConditionalFormat";
import type { EditProtection, EditProtectionContent } from "./EditProtection";
import type { RgbColor } from "./RgbColor";

export interface GridRangeProps {
  sheetId: number;
  startRowIndex: number;
  endRowIndex?: number;
  startColumnIndex?: number;
  endColumnIndex?: number;
}

export type BoundedGridRange = Required<GridRangeProps>;

export interface UsedGridRange {
  sheetId: number;
}

export type GridFetchRange = GridRangeProps | UsedGridRange;

export type SortOrder = "ASCENDING" | "DESCENDING";

export type FindReplaceScope =
  { range: GridRangeProps } | { sheetId: number } | { allSheets: true };

export interface FindReplaceTerms {
  find: string;
  replacement: string;
  matchCase?: boolean;
  matchEntireCell?: boolean;
  searchByRegex?: boolean;
  includeFormulas?: boolean;
}

export interface GridFetchOptions {
  includeProgrammaticFacts: boolean;
}

export interface SpreadsheetSnapshot {
  timeZone: string | null;
  sheets: SheetSnapshot[];
}

export interface SheetSnapshot {
  sheetGid: number;
  title: string | null;
  tables: TableSnapshot[] | undefined;
  gridBlocks: GridBlockSnapshot[] | undefined;
}

export interface SheetConditionalFormatSnapshot {
  sheetGid: number;
  rules: ConditionalFormatRule[];
}

export interface SheetEditProtectionSnapshot {
  sheetGid: number;
  protections: EditProtection[];
}

export interface TableSnapshot {
  tableId: string;
  name: string;
  startRowIndex: number;
  endRowIndex: number;
  startColumnIndex: number;
  endColumnIndex: number;
  columnProperties: TableColumnSnapshot[];
}

export interface TableColumnSnapshot {
  columnIndex: number;
  columnName?: string;
  columnType?: string;
  dataValidationValues: string[];
  dataValidationConditionType?: string;
}

// Google's Table column type enum, for writes; reads stay string so an unknown type resends unchanged.
export type TableColumnType =
  | "COLUMN_TYPE_UNSPECIFIED"
  | "DOUBLE"
  | "CURRENCY"
  | "PERCENT"
  | "DATE"
  | "TIME"
  | "DATE_TIME"
  | "TEXT"
  | "BOOLEAN"
  | "DROPDOWN"
  | "FILES_CHIP"
  | "PEOPLE_CHIP"
  | "FINANCE_CHIP"
  | "PLACE_CHIP"
  | "RATINGS_CHIP";

export interface GridBlockSnapshot {
  startColumn: number;
  startRow: number;
  columnCount: number;
  rows: GridRowSnapshot[];
}

export interface GridRowSnapshot {
  cells: (GridCellSnapshot | undefined)[];
}

export interface GridCellSnapshot {
  value: CellValue | "";
  isFormula: boolean;
  numberFormatType?: string;
  dataValidationConditionType?: string;
  backgroundColor?: RgbColor;
}

export type LocalWriteOperation =
  | AddSheetOperation
  | AddTableOperation
  | AppendRowsOperation
  | InsertColumnOperation
  | FillOperation
  | UpdateCellOperation
  | FindReplaceOperation
  | DeleteRowsOperation
  | SortOperation
  | AddConditionalFormatRuleOperation
  | DeleteConditionalFormatRuleOperation
  | AddProtectedRangeOperation
  | DeleteProtectedRangeOperation
  | UpdateSheetTitleOperation
  | UpdateTableNameOperation
  | UpdateTableColumnPropertiesOperation
  | AddCheckboxValidationOperation
  | OpaqueRawWriteOperation;

// Always at a given GID: Google refuses one another tab already holds (#74).
export interface AddSheetOperation {
  kind: "addSheet";
  sheetId: number;
  title: string;
  rowCount: number;
  columnCount: number;
}

// Carries no tableId: the adapter keys the Table by its name, which the floor matches it by.
export interface AddTableOperation {
  kind: "addTable";
  name: string;
  range: GridRangeProps;
  columnProperties: TableColumnPropertiesAdd[];
}

export interface TableColumnPropertiesAdd {
  columnIndex: number;
  columnName: string;
  columnType: TableColumnType;
}

export interface AppendRowsOperation {
  kind: "appendRows";
  sheetId: number;
  tableId: string;
  emptyRowCount: number;
}

export interface InsertColumnOperation {
  kind: "insertColumn";
  sheetId: number;
  startColumnIndex: number;
}

export interface FillOperation {
  kind: "fill";
  sheetId: number;
  colIndex: number;
  startRowIndex: number;
  endRowIndex: number;
  value?: CellValue;
  formula?: string;
  backgroundColor?: RgbColor;
}

export interface UpdateCellOperation {
  kind: "updateCell";
  sheetId: number;
  rowIndex: number;
  colIndex: number;
  value?: CellValue;
  formula?: string;
  backgroundColor?: RgbColor;
}

export interface FindReplaceOperation {
  kind: "findReplace";
  terms: FindReplaceTerms;
  scope: FindReplaceScope;
}

export interface DeleteRowsOperation {
  kind: "deleteRows";
  sheetId: number;
  startIndex: number;
  endIndex: number;
}

export interface SortOperation {
  kind: "sort";
  sheetId: number;
  startRowIndex: number;
  startColumnIndex: number;
  colIdxToSortBy: number;
  sortOrder: SortOrder;
}

export interface AddConditionalFormatRuleOperation {
  kind: "addConditionalFormatRule";
  index: number;
  rule: ModelableConditionalFormatRule;
}

export interface DeleteConditionalFormatRuleOperation {
  kind: "deleteConditionalFormatRule";
  sheetId: number;
  index: number;
}

export interface AddProtectedRangeOperation {
  kind: "addProtectedRange";
  protection: EditProtectionContent;
}

export interface DeleteProtectedRangeOperation {
  kind: "deleteProtectedRange";
  sheetId: number;
  protectedRangeId: number;
}

export interface UpdateSheetTitleOperation {
  kind: "updateSheetTitle";
  sheetId: number;
  title: string;
}

export interface UpdateTableNameOperation {
  kind: "updateTableName";
  tableId: string;
  name: string;
}

// Replaces the Table's whole column list, so it carries every column.
export interface UpdateTableColumnPropertiesOperation {
  kind: "updateTableColumnProperties";
  tableId: string;
  columnProperties: TableColumnPropertiesUpdate[];
}

export interface TableColumnPropertiesUpdate {
  columnIndex: number;
  columnName: string;
  columnType?: string;
}

export interface AddCheckboxValidationOperation {
  kind: "addCheckboxValidation";
  range: BoundedGridRange;
}

declare const opaqueRawRequest: unique symbol;
// Only the platform module can build or read one.
export interface OpaqueRawRequest {
  readonly [opaqueRawRequest]: true;
}

export interface OpaqueRawWriteOperation {
  kind: "raw";
  request: OpaqueRawRequest;
}

export interface RawSource {
  fetchSheetProperties(): SpreadsheetSnapshot;
  fetchTimeZone(): string | null;
  fetchGrid(
    gridRanges: GridFetchRange[],
    options: GridFetchOptions,
  ): SpreadsheetSnapshot;
  fetchConditionalFormatRules(): SheetConditionalFormatSnapshot[];
  fetchEditProtections(): SheetEditProtectionSnapshot[];
  flush(operations: LocalWriteOperation[]): void;
}

let installed: RawSource | undefined;

export function installRawSource(source: RawSource): void {
  installed = source;
}

export function hasInstalledRawSource(): boolean {
  return installed !== undefined;
}

export function installedRawSource(): RawSource {
  if (installed === undefined) {
    throw new Error(
      "RawSource has not been installed. The host must construct GoogleSheetsAPI before SpreadsheetRaw.init.",
    );
  }
  return installed;
}
