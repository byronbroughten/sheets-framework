import type {
  RowWrites,
  SheetStateRaw,
  SpreadsheetFetchQueueRaw,
  SpreadsheetWriteQueueRaw,
  TableFetchQueueRaw,
  TableStateRaw,
  TableWorkingStateRaw,
  TableWriteQueueRaw,
  TableWrites,
  WriteOperations,
} from "./StateRaw";

export const emptyStateRaw = {
  writeOperations(): WriteOperations {
    return {
      addSheet: [],
      addTable: [],
      appendRows: [],
      fillCell: [],
      deleteTableRows: [],
      sort: [],
      insertTableEndColumn: [],
      fillColumn: [],
      findReplace: [],
      deleteConditionalFormatRule: [],
      addConditionalFormatRule: [],
      deleteProtectedRange: [],
      addProtectedRange: [],
      renameSheet: [],
      renameTable: [],
      setTableColumnProperties: [],
      addCheckboxValidation: [],
      raw: [],
    };
  },
  spreadsheetFetchQueue(): SpreadsheetFetchQueueRaw {
    return { gridRanges: [] };
  },
  spreadsheetWriteQueue(): SpreadsheetWriteQueueRaw {
    return { operations: emptyStateRaw.writeOperations() };
  },
  tableWrites(): TableWrites {
    return {
      sort: undefined,
      insertTableEndColumnCount: 0,
      fillColumns: [],
      columnTypes: new Map(),
    };
  },
  rowWrites(): RowWrites {
    return { appendRow: false, deleteRow: false, fillCells: new Map() };
  },
  tableWriteQueue(): TableWriteQueueRaw {
    return {
      table: emptyStateRaw.tableWrites(),
      rows: new Map(),
      reservedRowIndexes: new Set(),
    };
  },
  tableFetchQueue(): TableFetchQueueRaw {
    return {
      toFinalize: {
        rows: new Set(),
        columns: new Set(),
        cells: new Map(),
      },
    };
  },
  tableState(sheetGid: number): TableStateRaw {
    return {
      sheetGid,
      properties: undefined,
      working: emptyTableWorkingState(),
      fetchQueue: emptyStateRaw.tableFetchQueue(),
      writeQueue: emptyStateRaw.tableWriteQueue(),
    };
  },
  sheetState(sheetGid: number): SheetStateRaw {
    return {
      working: {
        title: undefined,
        rowCount: undefined,
        conditionalFormats: { rules: undefined, isStale: false },
        editProtections: { protections: undefined, isStale: false },
      },
      fetchQueue: {
        gatherConditionalFormats: false,
        gatherEditProtections: false,
        gatherPlacementStrip: false,
      },
      tableBeforeProperties: emptyStateRaw.tableState(sheetGid),
    };
  },
};

function emptyTableWorkingState(): TableWorkingStateRaw {
  return {
    cellStateIsStale: false,
    hasFetchedColumnIds: false,
    isPrunedToSelection: false,
    rowStates: new Map(),
    columnStates: new Map(),
  };
}
