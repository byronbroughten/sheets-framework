import type {
  CellValue,
  CellValueName,
} from "../../00_Source/CellValues/cellValues";
import type { ConditionalFormatRule } from "../../00_Source/RawSource/ConditionalFormat";
import type { EditProtection } from "../../00_Source/RawSource/EditProtection";
import type {
  FillCellOperation,
  FindReplaceScope as BaseFindReplaceScope,
  FindReplaceTerms as BaseFindReplaceTerms,
  LocalWriteOperation,
  RawSource,
  TableColumnSnapshot,
  TableColumnType,
} from "../../00_Source/RawSource/RawSource";
import type { RgbColor } from "../../00_Source/RawSource/RgbColor";
import type {
  SheetColIndex,
  SheetRowIndex,
} from "../../00_Source/RawSource/SheetIndex";
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
  operations: WriteOperations;
}

// Built from a Table's growth or delete as the flush gathers it, so never queued itself.
type GatheredOnlyKind =
  | "appendDimension"
  | "insertRange"
  | "deleteRange"
  | "copyPaste";

// A key equals its operation's kind, so the compiler rejects one that matches none.
export type WriteOperations = {
  [KD in Exclude<LocalWriteOperation["kind"], GatheredOnlyKind>]: Extract<
    LocalWriteOperation,
    { kind: KD }
  >[];
} & { setTableColumnType: SetTableColumnTypeOperation[] };

// Queue-only: the sheet gathers each Table's into one setTableColumnProperties.
export interface SetTableColumnTypeOperation {
  kind: "setTableColumnType";
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
  sheet: SheetWrites;
  rows: Map<RowIndex, RowWrites>;
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
  startRowIndex: SheetRowIndex; // the header row
  endRowIndex: SheetRowIndex; // last row + 1
  startColumnIndex: SheetColIndex;
  endColumnIndex: SheetColIndex; // last column + 1
  columnProperties: TableColumnSnapshot[];
  rowIndexesAreStale: boolean;
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

export interface RowWrites {
  appendRow: boolean;
  deleteRow: boolean;
  // Values, not indexes, so a queued write never depends on fetched row state.
  fillCells: Map<ColIndex, CellFill>;
}
// One entry per cell, merged across writes, so a colour never cancels a value or a formula.
export interface CellFill<VN extends CellValueName = CellValueName> {
  value?: CellValue<VN>;
  formula?: string;
  backgroundColor?: RgbColor;
}
// The uniform cells a Table-end column insert writes.
export interface TableEndColumnUniformCells {
  columnId: string;
  header: string;
  colGroupName?: string;
}
export interface SheetWrites {
  sort: SortParameters | undefined;
  insertTableEndColumnCount: number;
  fillColumns: ColumnFill[];
}
// One contiguous run of a column's cells: value/colour as repeatCell, formula as pasteData.
export interface ColumnFill extends CellFill {
  colIndex: ColIndex;
  startRowIndex: number;
  // Snapshotted when queued, so a fill never reaches a row appended after it.
  endRowIndex: number;
}

export interface SheetWriteSortProps extends SortParameters {
  action: "sort";
}

export type FindReplaceScope = BaseFindReplaceScope;
export type FindReplaceTerms = BaseFindReplaceTerms;
export interface FindReplaceProps extends FindReplaceTerms {
  scope: FindReplaceScope;
}

export type AddedSheetCell = Required<
  Pick<FillCellOperation, "sheetId" | "rowIndex" | "colIndex">
> &
  ({ value: CellValue } | { formula: string });

export interface SheetWritePropsObj {
  sort: SheetWriteSortProps;
  insertTableEndColumn: { action: "insertTableEndColumn" };
  fillColumn: { action: "fillColumn" } & ColumnFill;
}
export type SheetWriteProps = SheetWritePropsObj[keyof SheetWritePropsObj];

export type RowWriteFillCellProps = {
  action: "fillCell";
  colIndex: ColIndex;
} & (
  { value: CellValue } | { formula: string } | { backgroundColor: RgbColor }
);
export type RowWriteProps =
  { action: "appendRow" | "deleteRow" } | RowWriteFillCellProps;

export type ColumnSpecifierRaw = ColIndex[] | "allColumns";
export type ColumnCount = number | "allFromStart";
export type RowCountRaw = number | "allFromStart";
