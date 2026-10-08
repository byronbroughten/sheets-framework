import type {
  RowWrites,
  SheetStateRaw,
  SheetWriteQueueRaw,
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
      appendTableRows: [],
      fillCell: [],
      deleteTableRows: [],
      sortTable: [],
      insertTableEndColumns: [],
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
      findReplaces: [],
      columnTypes: new Map(),
      checkboxCells: [],
      conditionalFormatRules: [],
      editProtections: [],
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
  sheetWriteQueue(): SheetWriteQueueRaw {
    return { appendedRowCount: 0, appendedColumnCount: 0 };
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
        columnCount: undefined,
        conditionalFormats: { rules: undefined, isStale: false },
        editProtections: { protections: undefined, isStale: false },
      },
      fetchQueue: {
        gatherConditionalFormats: false,
        gatherEditProtections: false,
        gatherHeaderZone: false,
      },
      writeQueue: emptyStateRaw.sheetWriteQueue(),
      tableBeforeProperties: emptyStateRaw.tableState(sheetGid),
      tablesBeforePropertiesById: new Map(),
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
