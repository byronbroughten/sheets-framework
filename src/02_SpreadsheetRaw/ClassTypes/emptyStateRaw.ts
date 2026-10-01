import type {
  RowWrites,
  SheetFetchQueueRaw,
  SheetStateRaw,
  SheetWorkingStateRaw,
  SheetWriteQueueRaw,
  SheetWrites,
  SpreadsheetFetchQueueRaw,
  SpreadsheetWriteQueueRaw,
  WriteOperations,
} from "./StateRaw";

export const emptyStateRaw = {
  writeOperations(): WriteOperations {
    return {
      addSheet: [],
      addTable: [],
      appendRows: [],
      fillCell: [],
      deleteRows: [],
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
      setTableColumnType: [],
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
  sheetWrites(): SheetWrites {
    return { sort: undefined, insertTableEndColumnCount: 0, fillColumns: [] };
  },
  rowWrites(): RowWrites {
    return { appendRow: false, deleteRow: false, fillCells: new Map() };
  },
  sheetWriteQueue(): SheetWriteQueueRaw {
    return {
      sheet: emptyStateRaw.sheetWrites(),
      rows: new Map(),
      reservedRowIndexes: new Set(),
    };
  },
  sheetState(): SheetStateRaw {
    return {
      working: emptySheetWorkingState(),
      fetchQueue: emptySheetFetchQueue(),
      writeQueue: emptyStateRaw.sheetWriteQueue(),
    };
  },
};

function emptySheetFetchQueue(): SheetFetchQueueRaw {
  return {
    gatherConditionalFormats: false,
    gatherEditProtections: false,
    toFinalize: {
      rows: new Set(),
      columns: new Set(),
      cells: new Map(),
    },
  };
}

function emptySheetWorkingState(): SheetWorkingStateRaw {
  return {
    title: undefined,
    knownTable: undefined,
    tables: [],
    hasExtraTables: false,
    cellStateIsStale: false,
    hasFetchedColumnIds: false,
    isPrunedToSelection: false,
    rowStates: new Map(),
    columnStates: new Map(),
    conditionalFormats: { rules: undefined, isStale: false },
    editProtections: { protections: undefined, isStale: false },
  };
}
