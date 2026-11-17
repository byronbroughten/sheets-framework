import type {
  CellValue,
  CellValueName,
} from "../../00_Source/CellValues/cellValues";
import type {
  ConditionalFormatDeclaration,
  ConditionalFormatRule,
} from "../../00_Source/RawSource/ConditionalFormat";
import type {
  EditProtection,
  EditProtectionContent,
} from "../../00_Source/RawSource/EditProtection";
import type {
  AddTableOperation,
  BoundedGridRange,
  CopyPasteOperation,
  CutPasteOperation,
  DeleteRangeOperation,
  FillCellOperation,
  FindReplaceScope,
  FindReplaceTerms,
  GridRangeProps,
  InsertRangeOperation,
  LocalWriteOperation,
  RawSource,
  SheetTabProperties,
  TableColumnSnapshot,
  TableColumnType,
  UpdateTableRangeOperation,
} from "../../00_Source/RawSource/RawSource";
import type { RgbColor } from "../../00_Source/RawSource/RgbColor";
import type {
  SheetColIndex,
  SheetRowIndex,
} from "../../00_Source/RawSource/SheetIndex";
import type { ColumnTypeBaseline } from "../../01_SpreadsheetSchema/columnBaseline";
import type { IdPrefixFromTitleProps } from "../../01_SpreadsheetSchema/idPrefixes";

export interface StateRaw {
  allSheetPropertiesAreFetched: boolean;
  // The Tables the last grid fetch gathered for, kept past its queue reset for the placement check.
  usedTableIds: Set<string>;
  timeZone: string | undefined;
  // The batches this run sent that Google took; an empty flush sends none.
  landedFlushCount: number;
  rawSource: RawSource;
  fetchQueue: SpreadsheetFetchQueueRaw;
  writeQueue: SpreadsheetWriteQueueRaw;
  sheets: SheetsStateRaw;
  tables: TablesStateRaw;
}

export interface SpreadsheetFetchQueueRaw {
  gridRanges: GridRangeProps[];
}

export interface SpreadsheetWriteQueueRaw {
  operations: WriteOperations;
}

// Built from a Table's growth, column insert, delete or move as the flush gathers it, so never queued itself.
type GatheredOnlyKind =
  | "appendDimension"
  | "insertRange"
  | "updateTableRange"
  | "deleteRange"
  | "copyPaste"
  | "cutPaste";

// A key equals its operation's kind, so the compiler rejects one that matches none.
export type WriteOperations = {
  [KD in Exclude<LocalWriteOperation["kind"], GatheredOnlyKind>]: Extract<
    LocalWriteOperation,
    { kind: KD }
  >[];
} & {
  appendTableRows: AppendTableRows[];
  insertTableEndColumns: InsertTableEndColumns[];
  changeTableColumns: ChangeTableColumns[];
};

// One Table's growth, expanded at gathering; the flusher sends a sheet's growths bottom-up.
export interface AppendTableRows {
  tableId: string;
  // Where the rows go in, in the layout before any of the batch's growth.
  newRows: BoundedGridRange;
  operations: TableGrowthOperation[];
}

export type TableGrowthOperation =
  InsertRangeOperation | UpdateTableRangeOperation | CopyPasteOperation;

// One Table's column inserts, expanded at gathering; the flusher sends a sheet's inserts right to left.
export interface InsertTableEndColumns {
  tableId: string;
  // From the column ID row to the last row, in the layout before any of the batch's column inserts.
  newColumns: BoundedGridRange;
  operations: TableColumnInsertOperation[];
}

export type TableColumnInsertOperation =
  InsertRangeOperation | UpdateTableRangeOperation;

// One Table's column inserts and deletes inside it, in queue order; the flusher sends a sheet's Tables right to left.
export interface ChangeTableColumns {
  tableId: string;
  sheetId: number;
  // In the layout before any of the batch's column changes.
  startColumnIndex: SheetColIndex;
  operations: TableColumnChangeOperation[];
}

export type TableColumnChangeOperation =
  InsertRangeOperation | DeleteRangeOperation | CutPasteOperation;

export type SheetsStateRaw = Map<SheetId, SheetStateRaw>;

export interface SheetStateRaw {
  working: SheetWorkingStateRaw;
  fetchQueue: SheetFetchQueueRaw;
  writeQueue: SheetWriteQueueRaw;
  // What is queued on a Table reached by its ID before its properties are fetched, keyed by that ID.
  tablesBeforePropertiesById: Map<string, TableStateRaw>;
}

export interface SheetWorkingStateRaw {
  title: string | undefined;
  rowCount: number | undefined;
  columnCount: number | undefined;
  tabProperties: SheetTabProperties | undefined;
  // So a column insert gathered later can judge the sheet's Tables from the zone already fetched.
  hasFetchedHeaderZone: boolean;
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
  // So the sheet's recorded Tables are judged on the fetch it rides.
  gatherHeaderZone: boolean;
}

// Summed over every growth or column insert on the sheet, and sent as one appendDimension each.
export interface SheetWriteQueueRaw {
  appendedRowCount: number;
  appendedColumnCount: number;
}

export type TablesStateRaw = Map<TableId, TableStateRaw>;

export interface TableStateRaw {
  sheetGid: number;
  // Absent until a fetch brings the Table's properties.
  properties: TablePropertiesRaw | undefined;
  working: TableWorkingStateRaw;
  fetchQueue: TableFetchQueueRaw;
  writeQueue: TableWriteQueueRaw;
}

export interface TablePropertiesRaw {
  tableId: string;
  name: string;
  startRowIndex: SheetRowIndex; // the header row
  endRowIndex: SheetRowIndex; // last row + 1
  startColumnIndex: SheetColIndex;
  endColumnIndex: SheetColIndex; // last column + 1
  columnProperties: TableColumnSnapshot[];
  rowIndexesAreStale: boolean;
}

export interface TableWorkingStateRaw {
  // A findReplace matches by content, so what it changed is unknowable locally.
  cellStateIsStale: boolean;
  hasFetchedColumnIds: boolean;
  isPrunedToSelection: boolean;
  rowStates: RowStatesRaw;
  columnStates: ColumnStatesRaw;
}

export interface TableFetchQueueRaw {
  toFinalize: TableFinalizeQueueRaw;
}

export interface TableFinalizeQueueRaw {
  rows: Set<RowIndex>;
  columns: Set<ColIndex>;
  cells: Map<RowIndex, Set<ColIndex>>;
}

export interface TableWriteQueueRaw {
  table: TableWrites;
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
  sampledFacts?: SampledFactsRaw;
  validationValues?: string[];
  validationConditionType?: string;
  // Absent for a column left on Automatic, which is what makes it "untyped".
  columnType?: string;
}
export interface SampledFactsRaw {
  isFormula: boolean;
  numberFormatType: string | undefined;
  dataValidationConditionType: string | undefined;
  topValue: CellValue;
  topFormula: string | undefined;
}

type SheetId = number;
type TableId = string;
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
// The head cells a column insert writes.
export interface NewColumnHeadCells {
  columnId: string;
  header: string;
  groupHeading1?: string;
}
// Its index is in the layout the Table's earlier queued changes leave.
export type TableColumnChange =
  | { kind: "insertColumn"; colIndex: ColIndex; headCells: NewColumnHeadCells }
  | { kind: "deleteColumn"; colIndex: ColIndex }
  | ({ kind: "moveColumns" } & ColumnMoveSpan);
// A block of count columns from colIndex, landing before beforeColIndex as the layout stands before the move.
export interface ColumnMoveSpan {
  colIndex: ColIndex;
  count: number;
  beforeColIndex: ColIndex;
}
// A bare type keeps the column's fetched rule; a baseline sends its own rule, and Automatic for a null type.
export type QueuedColumnType = TableColumnType | ColumnTypeBaseline;
export interface TableWrites {
  sort: SortParameters | undefined;
  insertTableEndColumnCount: number;
  // Sent in queue order, so each index is read after the changes before it.
  columnChanges: TableColumnChange[];
  fillColumns: ColumnFill[];
  findReplaces: TableFindReplace[];
  // Gathered into the Table's one setTableColumnProperties.
  columnTypes: Map<ColIndex, QueuedColumnType>;
  // Raw can't tell a formula column, so an append's caller names the ones growth copies.
  appendedFormulaColIndexes: Set<ColIndex>;
  checkboxCells: TableCell[];
  conditionalFormatRules: TableConditionalFormatRule[];
  editProtections: TableEditProtection[];
  deletesTable: boolean;
}
// Table-relative until gathering; an absent bound stays open, and an absent start row is the sheet's top.
export interface TableGridRange {
  startRowIndex?: RowIndex;
  endRowIndex?: RowIndex;
  startColIndex?: ColIndex;
  endColIndex?: ColIndex;
}
export type BoundedTableGridRange = Required<TableGridRange>;
export interface TableCell {
  rowIndex: RowIndex;
  colIndex: ColIndex;
}
export interface TableConditionalFormatRule extends ConditionalFormatDeclaration {
  range: TableGridRange;
}
export interface TableEditProtection extends Omit<
  EditProtectionContent,
  "range"
> {
  range: TableGridRange;
}
// One contiguous run of a column's cells: value/colour as repeatCell, formula as pasteData.
export interface ColumnFill extends CellFill {
  colIndex: ColIndex;
  startRowIndex: number;
  // Snapshotted when queued, so a fill never reaches a row appended after it.
  endRowIndex: number;
}

export interface TableFindReplace {
  terms: FindReplaceTerms;
  startColIndex: ColIndex;
  endColIndex: ColIndex;
}

export interface TableWriteSortProps extends SortParameters {
  action: "sort";
}

export interface FindReplaceProps extends FindReplaceTerms {
  scope: FindReplaceScope;
}

// A caller supplies a tableId only to recreate a Table at its recorded one.
export type AddTableProps = Omit<AddTableOperation, "kind" | "tableId"> &
  Partial<Pick<AddTableOperation, "tableId">>;

export type AddedSheetCell = Required<
  Pick<FillCellOperation, "sheetId" | "rowIndex" | "colIndex">
> &
  ({ value: CellValue } | { formula: string });

// A tab the same flush adds, or right of the rightmost Table on a tab already there, `gap` empty columns between.
export type CreatedTablePlacement =
  | { kind: "newTab"; sheetGid: number; title: string }
  | { kind: "besideTables"; sheetGid: number; gap?: 0 | 1 };

export interface CreatedTableColumn {
  header: string;
  columnId?: string;
  columnType?: ColumnTypeBaseline;
  // Already in the names on the sheet; filled into every body row.
  formula?: string;
}

// A fixed prefix, or the prefixes a prefix drawn from the Table's name must skip.
export type CreatedTableIdPrefix =
  string | Omit<IdPrefixFromTitleProps, "title">;

export interface CreateTableProps {
  name: string;
  placement: CreatedTablePlacement;
  idPrefix: CreatedTableIdPrefix;
  columns: CreatedTableColumn[];
  tableId?: string;
  // Each row's values in column order; without any, the body is one empty row.
  seedRows?: CellValue[][];
}

export interface CreatedTable {
  tableId: string;
  idPrefix: string;
  columnIds: string[];
  // Header row through the last body row, where the addTable puts it.
  range: BoundedGridRange;
}

export interface TableWritePropsObj {
  sort: TableWriteSortProps;
  insertTableEndColumn: { action: "insertTableEndColumn" };
  insertColumn: {
    action: "insertColumn";
    colIndex: ColIndex;
    headCells: NewColumnHeadCells;
  };
  deleteColumn: { action: "deleteColumn"; colIndex: ColIndex };
  moveColumns: { action: "moveColumns" } & ColumnMoveSpan;
  fillColumn: { action: "fillColumn" } & ColumnFill;
  findReplace: { action: "findReplace" } & TableFindReplace;
  updateColumnType: {
    action: "updateColumnType";
    colIndex: ColIndex;
    columnType: QueuedColumnType;
  };
  addCheckboxValidation: { action: "addCheckboxValidation" } & TableCell;
  copyAppendedFormulas: {
    action: "copyAppendedFormulas";
    colIndexes: readonly ColIndex[];
  };
  addConditionalFormatRule: {
    action: "addConditionalFormatRule";
  } & TableConditionalFormatRule;
  addEditProtection: { action: "addEditProtection" } & TableEditProtection;
  deleteTable: { action: "deleteTable" };
}
export type TableWriteProps = TableWritePropsObj[keyof TableWritePropsObj];

export type RowWriteFillCellProps = {
  action: "fillCell";
  colIndex: ColIndex;
} & (
  { value: CellValue } | { formula: string } | { backgroundColor: RgbColor }
);
export type RowWriteProps =
  { action: "appendRow" | "deleteRow" } | RowWriteFillCellProps;
