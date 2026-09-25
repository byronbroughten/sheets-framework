import type {
  CellValue,
  CellValueName,
} from "../../00_Source/CellValues/cellValues";
import type { ConditionalFormatRule } from "../../00_Source/RawSource/ConditionalFormat";
import type { EditProtection } from "../../00_Source/RawSource/EditProtection";
import type {
  AddCheckboxValidationOperation,
  AddConditionalFormatRuleOperation,
  AddProtectedRangeOperation,
  AddSheetOperation,
  AddTableOperation,
  AppendRowsOperation,
  DeleteConditionalFormatRuleOperation,
  DeleteProtectedRangeOperation,
  DeleteRowsOperation,
  FillOperation,
  FindReplaceOperation,
  FindReplaceScope as BaseFindReplaceScope,
  FindReplaceTerms as BaseFindReplaceTerms,
  InsertColumnOperation,
  OpaqueRawWriteOperation,
  RawSource,
  SortOperation,
  TableColumnSnapshot,
  TableColumnType,
  UpdateCellOperation,
  UpdateSheetTitleOperation,
  UpdateTableColumnPropertiesOperation,
  UpdateTableNameOperation,
} from "../../00_Source/RawSource/RawSource";
import type { RgbColor } from "../../00_Source/RawSource/RgbColor";
import type { GridRangeProps } from "./AccessorsRaw";

export interface StateRaw {
  allSheetPropertiesAreFetched: boolean;
  timeZone: string | undefined;
  rawSource: RawSource;
  fetchQueue: SpreadsheetFetchQueueRaw;
  writeQueue: SpreadsheetWriteQueueRaw;
  sheets: SheetsStateRaw;
}

export interface SpreadsheetFetchQueueRaw {
  gridRanges: GridRangeProps[];
}

export interface SpreadsheetWriteQueueRaw {
  updateRequests: UpdateRequests;
}

export interface UpdateRequests {
  addSheet: AddSheetOperation[];
  addTable: AddTableOperation[];
  append: AppendRowsOperation[];
  update: UpdateCellOperation[];
  delete: DeleteRowsOperation[];
  sort: SortOperation[];
  insertColumn: InsertColumnOperation[];
  fill: FillOperation[];
  findReplace: FindReplaceOperation[];
  deleteConditionalFormat: DeleteConditionalFormatRuleOperation[];
  addConditionalFormat: AddConditionalFormatRuleOperation[];
  deleteProtectedRange: DeleteProtectedRangeOperation[];
  addProtectedRange: AddProtectedRangeOperation[];
  updateSheetTitle: UpdateSheetTitleOperation[];
  updateTableName: UpdateTableNameOperation[];
  updateTableColumnType: UpdateTableColumnTypeOperation[];
  updateTableColumnProperties: UpdateTableColumnPropertiesOperation[];
  addCheckboxValidation: AddCheckboxValidationOperation[];
  raw: OpaqueRawWriteOperation[];
}

// Queue-only: the sheet gathers each Table's into one updateTableColumnProperties.
export interface UpdateTableColumnTypeOperation {
  kind: "updateTableColumnType";
  sheetId: number;
  tableId: string;
  columnIndex: number;
  columnType: TableColumnType;
}

export type SheetsStateRaw = Map<SheetId, SheetStateRaw>;

export interface SheetStateRaw {
  working: SheetWorkingStateRaw;
  fetchQueue: SheetFetchQueueRaw;
  writeQueue: SheetWriteQueueRaw;
}

export interface SheetWorkingStateRaw {
  title: string | undefined;
  knownTable: KnownTableRaw | undefined;
  tables: TableIdentityRaw[];
  hasExtraTables: boolean;
  // A findReplace matches by content, so what it changed is unknowable locally.
  cellStateIsStale: boolean;
  hasFetchedColumnIds: boolean;
  isPrunedToSelection: boolean;
  rowStates: RowStatesRaw;
  columnStates: ColumnStatesRaw;
  conditionalFormats: ConditionalFormatsStateRaw;
  editProtections: EditProtectionsStateRaw;
}

export interface ConditionalFormatsStateRaw {
  rules: ConditionalFormatRule[] | undefined;
  isStale: boolean;
}

export interface EditProtectionsStateRaw {
  protections: EditProtection[] | undefined;
  isStale: boolean;
}

export interface SheetFetchQueueRaw {
  gatherConditionalFormats: boolean;
  gatherEditProtections: boolean;
  toFinalize: SheetFinalizeQueueRaw;
}

export interface SheetFinalizeQueueRaw {
  rows: Set<RowIndex>;
  columns: Set<ColIndex>;
  cells: Map<RowIndex, Set<ColIndex>>;
}

export interface SheetWriteQueueRaw {
  sheet: SheetChangesToSave;
  rows: Map<RowIndex, RowChangesToSave>;
  // A row an append has handed out, so a second append can't reuse it.
  reservedRowIndexes: Set<RowIndex>;
}

export type RowStatesRaw = Map<RowIndex, RowStateRaw>;
export type RowStateRaw = Map<ColIndex, CellStateRaw>;
export interface CellStateRaw {
  value: CellValue;
}

export type ColumnStatesRaw = Map<ColIndex, ColumnStateRaw>;
export interface ColumnStateRaw {
  activeFacts?: ActiveFactsRaw;
  validationValues?: string[];
  validationConditionType?: string;
  // Absent for a column left on Automatic, which is what makes it "untyped".
  columnType?: string;
}
export interface ActiveFactsRaw {
  isFormula: boolean;
  numberFormatType: string | undefined;
  dataValidationConditionType: string | undefined;
  topValue: CellValue;
}

export interface KnownTableRaw {
  tableId: string;
  name: string;
  startRowIndex: number; // tableHeaderRowIndex
  endRowIndex: number; // lastRowIndex + 1
  startColumnIndex: number;
  endColumnIndex: number; // lastColumnIndex + 1
  columnProperties: TableColumnSnapshot[];
  rowIndexesAreStale: boolean;
  firstStaleColIndex: number | undefined;
}

export interface TableIdentityRaw {
  tableId: string;
  name: string;
}

type SheetId = number;
type RowIndex = number;
type ColIndex = number;

export interface SortParameters {
  colIdxToSortBy: number;
  sortOrder: "ASCENDING" | "DESCENDING";
}

export interface RowChangesToSave {
  append: boolean;
  delete: boolean;
  // Values, not indexes, so a queued write never depends on fetched row state.
  update: Map<ColIndex, RowCellChange>;
}
// One entry per cell, merged across writes, so a colour never cancels a value or a formula.
export interface RowCellChange<VN extends CellValueName = CellValueName> {
  value?: CellValue<VN>;
  formula?: string;
  backgroundColor?: RgbColor;
}
export interface SheetChangesToSave {
  sort: SortParameters | undefined;
  insertColumn: SheetChangePropsObj["insertColumn"][];
  fills: ColumnFill[];
}
// One contiguous run of a column's cells: value/colour as repeatCell, formula as pasteData.
export interface ColumnFill extends RowCellChange {
  colIndex: ColIndex;
  startRowIndex: number;
  // Snapshotted when queued, so a fill never reaches a row appended after it.
  endRowIndex: number;
}

export interface SheetChangeSortProps extends SortParameters {
  action: "sort";
}

export type FindReplaceScope = BaseFindReplaceScope;
export type FindReplaceTerms = BaseFindReplaceTerms;
export interface FindReplaceProps extends FindReplaceTerms {
  scope: FindReplaceScope;
}

export type AddedSheetCell = Required<
  Pick<UpdateCellOperation, "sheetId" | "rowIndex" | "colIndex">
> &
  ({ value: CellValue } | { formula: string });

export interface SheetChangePropsObj {
  sort: SheetChangeSortProps;
  insertColumn: {
    action: "insertColumn";
    startColumnIndex: number;
  };
  fill: { action: "fill" } & ColumnFill;
}
export type SheetChangeProps = SheetChangePropsObj[keyof SheetChangePropsObj];

export type RowChangeUpdateProps = {
  action: "update";
  colIndex: ColIndex;
} & (
  { value: CellValue } | { formula: string } | { backgroundColor: RgbColor }
);
export type RowChangeProps =
  { action: "append" | "delete" } | RowChangeUpdateProps;

export type ColumnSpecifierRaw = ColIndex[] | "allColumns";
export type ColumnCount = number | "allFromStart";
export type RowCountRaw = number | "allFromStart";
