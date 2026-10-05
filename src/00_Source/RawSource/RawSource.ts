import type { CellValue } from "../CellValues/cellValues";
import type {
  ConditionalFormatRule,
  ModelableConditionalFormatRule,
} from "./ConditionalFormat";
import type { EditProtection, EditProtectionContent } from "./EditProtection";
import type { RgbColor } from "./RgbColor";
import type { SheetColIndex, SheetRowIndex } from "./SheetIndex";

export interface GridRangeProps {
  sheetId: number;
  startRowIndex: SheetRowIndex;
  endRowIndex?: SheetRowIndex;
  startColumnIndex?: SheetColIndex;
  endColumnIndex?: SheetColIndex;
}

export type BoundedGridRange = Required<GridRangeProps>;

export interface UsedGridRange {
  sheetId: number;
}

export type GridFetchRange = GridRangeProps | UsedGridRange;

export type SortOrder = "ASCENDING" | "DESCENDING";

export type GridDimension = "ROWS" | "COLUMNS";

// No sheet scope: it would reach a Table's head rows (sheets-framework#54).
export type FindReplaceScope = { range: GridRangeProps } | { allSheets: true };

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
  rowCount: number | undefined;
  columnCount: number | undefined;
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
  startRowIndex: SheetRowIndex;
  endRowIndex: SheetRowIndex;
  startColumnIndex: SheetColIndex;
  endColumnIndex: SheetColIndex;
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
  startColumn: SheetColIndex;
  startRow: SheetRowIndex;
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
  | AppendDimensionOperation
  | InsertRangeOperation
  | UpdateTableRangeOperation
  | DeleteRangeOperation
  | CopyPasteOperation
  | FillColumnOperation
  | FillCellOperation
  | FindReplaceOperation
  | DeleteTableRowsOperation
  | SortTableOperation
  | AddConditionalFormatRuleOperation
  | DeleteConditionalFormatRuleOperation
  | AddProtectedRangeOperation
  | DeleteProtectedRangeOperation
  | RenameSheetOperation
  | RenameTableOperation
  | SetTableColumnPropertiesOperation
  | AddCheckboxValidationOperation
  | RawWriteOperation;

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

// Adds rows or columns past the grid's edge, so it can't split a Table (sheets-framework#59).
export interface AppendDimensionOperation {
  kind: "appendDimension";
  sheetId: number;
  dimension: GridDimension;
  addedCount: number;
}

// Shifts only the range's own cells, so a Table beside it is spared (sheets-framework#53).
export interface InsertRangeOperation {
  kind: "insertRange";
  range: BoundedGridRange;
  shiftDimension: GridDimension;
}

// An insertRange below a Table leaves its range as it was, so growth widens it (sheets-framework#59).
export interface UpdateTableRangeOperation {
  kind: "updateTableRange";
  tableId: string;
  range: BoundedGridRange;
}

export interface DeleteRangeOperation {
  kind: "deleteRange";
  range: BoundedGridRange;
  shiftDimension: GridDimension;
}

export type CopyPasteType = "PASTE_FORMAT" | "PASTE_DATA_VALIDATION";

// One source row tiles over every destination row (sheets-framework#61).
export interface CopyPasteOperation {
  kind: "copyPaste";
  source: BoundedGridRange;
  destination: BoundedGridRange;
  pasteType: CopyPasteType;
}

export interface FillColumnOperation {
  kind: "fillColumn";
  sheetId: number;
  colIndex: SheetColIndex;
  startRowIndex: SheetRowIndex;
  endRowIndex: SheetRowIndex;
  value?: CellValue;
  formula?: string;
  backgroundColor?: RgbColor;
}

export interface FillCellOperation {
  kind: "fillCell";
  sheetId: number;
  rowIndex: SheetRowIndex;
  colIndex: SheetColIndex;
  value?: CellValue;
  formula?: string;
  backgroundColor?: RgbColor;
}

export interface FindReplaceOperation {
  kind: "findReplace";
  terms: FindReplaceTerms;
  scope: FindReplaceScope;
}

// Over the Table's columns only, so a side-by-side neighbour is never cut (sheets-framework#52).
export interface DeleteTableRowsOperation {
  kind: "deleteTableRows";
  range: BoundedGridRange;
}

// The Table's body only, so its head rows and neighbours never move (sheets-framework#95).
export interface SortTableOperation {
  kind: "sortTable";
  range: BoundedGridRange;
  colIdxToSortBy: SheetColIndex;
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

export interface RenameSheetOperation {
  kind: "renameSheet";
  sheetId: number;
  title: string;
}

export interface RenameTableOperation {
  kind: "renameTable";
  tableId: string;
  name: string;
}

// Replaces the Table's whole column list, so it carries every column.
export interface SetTableColumnPropertiesOperation {
  kind: "setTableColumnProperties";
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

export interface RawWriteOperation {
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
