import { Val } from "@byronbroughten/utils/val";

import type { DropdownStyleReset } from "../00_Source/RawSource/RawSource";
import { installedConfigs } from "../01_SpreadsheetSchema/configReaders/configRegister";
import { type IdNameMapGeneric } from "../01_SpreadsheetSchema/configReaders/idNameMapTypes";
import { type PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState";
import {
  type ForcedTrait,
  type SyncRunFrame,
} from "../01_SpreadsheetSchema/configReaders/syncRunResult";
import { getTableTraitByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { dimensionIds } from "../01_SpreadsheetSchema/dimensionIds";
import {
  layoutGid,
  layoutTableId,
  managedLayoutTableNames,
} from "../02_SpreadsheetRaw/spreadsheetRawTestSupport";
import { expectedSheetLayout } from "../testSupport/expectedSheetLayout";
import {
  type FakeSheetProperties,
  type FakeSheetsService,
  type FakeSheetsServiceOptions,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { a1Notation } from "../testSupport/fakeSheetsService/a1Notation";
import {
  type FakeBodyRow,
  type FakeTableColumn,
  fakeTableSheet,
  type FakeTableSheetProps,
} from "../testSupport/fakeSheetsService/fakeTableSheet";
import {
  type ConfigRegeneration,
  SyncCoordinator,
  type SyncRun,
} from "./SyncCoordinator";

interface ReadGridRange {
  sheetId?: number;
  startColumnIndex?: number;
  endColumnIndex?: number;
}

export const { columnConfigs } = installedConfigs();
export const itemGid = getTableTraitByName("item", "sheetGid");
export const itemColumns = columnConfigs.item;
export type ItemColumnName = keyof typeof itemColumns;

// By absolute column index.
type ItemColumnPassThrough = Pick<
  FakeTableSheetProps<string>,
  "columnTypes" | "columnValidationConditionTypes" | "columnValidationValues"
>;

interface ItemSheetOptions extends ItemColumnPassThrough {
  columns?: Record<string, FakeTableColumn>; // Replaces a configured column by name, or adds one after them.
  startRowIndex?: number;
  name?: string; // The Table's Sheets name; "Item" when unset.
  columnNames?: string[]; // The sheet's column order; the configured order when unset.
  bodyRows?: FakeBodyRow<string>[]; // One row with an ID when unset.
  actionRowCells?: FakeBodyRow<string>; // A Fill Row IDs checkbox above ID when unset.
}

export const itemRowId = "r:itm:fixture";

// The dev Item Table as the previous configs record it, with its ID filled and its checkbox in place, so every run starts quiet.
export function itemSheet(options: ItemSheetOptions = {}): FakeSheetProperties {
  const sheetColumns: Record<string, FakeTableColumn> = {
    ...itemColumns,
    ...options.columns,
  };
  return fakeTableSheet.build({
    sheetId: itemGid,
    title: "Item",
    tableId: "item",
    name: options.name ?? "Item",
    columnConfigs: sheetColumns,
    columnNames: options.columnNames ?? Object.keys(sheetColumns),
    bodyRows: options.bodyRows ?? [{ id: itemRowId }],
    actionRowCells: options.actionRowCells ?? { id: false },
    startRowIndex: options.startRowIndex,
    columnTypes: options.columnTypes,
    columnValidationConditionTypes: options.columnValidationConditionTypes,
    columnValidationValues: options.columnValidationValues,
  });
}

export function seedFixture(
  options: {
    item?: ItemSheetOptions;
    extraSheets?: FakeSheetProperties[];
    service?: Omit<FakeSheetsServiceOptions, "sheets">;
  } = {},
): FakeSheetsService {
  return stubSheetsService({
    ...options.service,
    sheets: [itemSheet(options.item), ...(options.extraSheets ?? [])],
  });
}

// Keyed by header, so a fixture names each column by what its sheet shows.
export function columnsByHeader(
  headers: readonly string[],
  columnIds: readonly string[] = [],
): Record<string, FakeTableColumn> {
  return Object.fromEntries(
    headers.map((header, colIndex) => [
      header,
      { columnId: columnIds[colIndex] ?? "", header },
    ]),
  );
}

// A Table the previous configs don't hold, at a GID they don't hold either.
export const notesGid = 999001;

export function notesSheet(): FakeSheetProperties {
  return fakeTableSheet.build({
    sheetId: notesGid,
    title: "Notes",
    tableId: "notes",
    name: "Desk Notes",
    columnConfigs: columnsByHeader(["Note"]),
    columnNames: ["Note"],
    bodyRows: [{}],
  });
}

interface TableConfigsEntryFixture {
  tableKey: string;
  tableId: string;
  tableName: string;
  sheetGid: number;
  idPrefix: string;
  hasIdColumn: boolean;
}

// Every fixture Table sits at the expected origin and carries a "Name" header.
export function tableConfigsEntry({
  tableKey,
  tableId,
  tableName,
  sheetGid,
  idPrefix,
  hasIdColumn,
}: TableConfigsEntryFixture): string {
  return `"${tableKey}": { "tableId": "${tableId}", "tableName": "${tableName}", "sheetGid": ${sheetGid}, "idPrefix": "${idPrefix}", "hasIdColumn": ${hasIdColumn}, "hasNameColumn": true }`;
}

export const frame: SyncRunFrame = {
  target: "dev",
  command: "npm run dev:sync",
  generatedDir: "dev/generated",
  alterationsPath: "dev/alterations.ts",
};

// The dev Item Table as the previous ID-name map records it.
export function previousItemMap(): IdNameMapGeneric {
  return {
    tables: {
      item: {
        name: "item",
        sheetGid: itemGid,
        columns: Object.fromEntries(
          Object.entries(itemColumns).map(([columnName, { columnId }]) => [
            columnName,
            columnId,
          ]),
        ),
      },
    },
    tabs: {
      [String(itemGid)]: { name: "item", title: "Item", tableIds: ["item"] },
    },
  };
}

const logGid = getTableTraitByName("log", "sheetGid");
const logColumns = columnConfigs.log;
const logColumnNames = Object.keys(logColumns);
const logFormulaColumn = Val.assert(logColumnNames.at(-1), "a Log column");
const logFormulaLetter = a1Notation.columnLetters(
  logColumnNames.indexOf(logFormulaColumn),
);
// The top body cell of Log's formula column, as an outside-references check lists it.
export const logFormulaCell = `Log!${logFormulaLetter}5:${logFormulaLetter}5`;

// The dev Log Table with one formula in its last column; a brace decompile matches each Table's live name.
function logSheet(formula: string): FakeSheetProperties {
  return fakeTableSheet.build({
    sheetId: logGid,
    title: "Log",
    tableId: getTableTraitByName("log", "tableId"),
    name: getTableTraitByName("log", "tableName"),
    columnConfigs: logColumns,
    columnNames: logColumnNames,
    bodyRows: [{ [logFormulaColumn]: { value: 1, isFormula: true, formula } }],
  });
}

function previousItemAndLogMap(): IdNameMapGeneric {
  const previous = previousItemMap();
  return {
    tables: {
      ...previous.tables,
      log: {
        name: "log",
        sheetGid: logGid,
        columns: Object.fromEntries(
          Object.entries(logColumns).map(([columnName, { columnId }]) => [
            columnName,
            columnId,
          ]),
        ),
      },
    },
    tabs: {
      ...previous.tabs,
      [String(logGid)]: {
        name: "log",
        title: "Log",
        tableIds: [getTableTraitByName("log", "tableId")],
      },
    },
  };
}

export const uncheckedOutsideFormulasNotice = {
  kind: "text",
  text: "A deletion's outside formulas are checked only in the API-access Tables' formula columns; a formula anywhere else that reads a deleted Table or tab isn't listed.",
};

interface BesideLogProps {
  alterations: PendingAlterations;
  logFormula: string;
  // A Table without API access beside Item on its tab, by its live Sheets name.
  unmanagedItemTabTable?: string;
}

const unmanagedTableId = "unmanaged";

// A dry run on Item and Log, Log's formula column holding logFormula; Item carries the name the previous configs record, which a live formula names it by.
export function dryRunBesideLog({
  alterations,
  logFormula,
  unmanagedItemTabTable,
}: BesideLogProps): SyncRun {
  let itemTab = itemSheet({ name: getTableTraitByName("item", "tableName") });
  const previousIdNameMap = previousItemAndLogMap();
  if (unmanagedItemTabTable !== undefined) {
    itemTab = withUnmanagedTable(itemTab, unmanagedItemTabTable);
    addUnmanagedTable(previousIdNameMap);
  }
  stubSheetsService({ sheets: [itemTab, logSheet(logFormula)] });
  return SyncCoordinator.init().run({
    isSending: false,
    intendedState: { alterations },
    previousIdNameMap,
    forcedTraits: [],
    frame,
  });
}

// Right of Item, past its columns, so the two Tables share no cell.
function withUnmanagedTable(
  item: FakeSheetProperties,
  name: string,
): FakeSheetProperties {
  const unmanaged = fakeTableSheet.build({
    sheetId: itemGid,
    title: "Item",
    tableId: unmanagedTableId,
    name,
    columnConfigs: columnsByHeader(["Amount"]),
    columnNames: ["Amount"],
    bodyRows: [{ Amount: 1 }],
    startColumnIndex: Object.keys(itemColumns).length + 2,
  });
  const itemRows = item.rows ?? [];
  const unmanagedRows = unmanaged.rows ?? [];
  const rowCount = Math.max(itemRows.length, unmanagedRows.length);
  const rows = Array.from({ length: rowCount }, (_, rowIndex) => {
    const itemRow = itemRows[rowIndex] ?? [];
    const unmanagedRow = unmanagedRows[rowIndex] ?? [];
    const colCount = Math.max(itemRow.length, unmanagedRow.length);
    return Array.from(
      { length: colCount },
      (_, colIndex) => itemRow[colIndex] ?? unmanagedRow[colIndex] ?? null,
    );
  });
  return {
    ...item,
    rows,
    tables: [...(item.tables ?? []), ...(unmanaged.tables ?? [])],
  };
}

// Its map key isn't its live name, as for any Table without API access the sync's titleToName reshapes.
function addUnmanagedTable(map: IdNameMapGeneric): void {
  map.tables[unmanagedTableId] = { name: unmanagedTableId, sheetGid: itemGid };
  const itemTab = Val.assert(map.tabs[String(itemGid)], "the item tab");
  itemTab.tableIds.push(unmanagedTableId);
}

export interface RunOptions {
  isSending: boolean;
  alterations?: PendingAlterations;
  forcedTraits?: ForcedTrait[];
  dropdownStyleReset?: DropdownStyleReset;
}

export function runItem({
  isSending,
  alterations = {},
  forcedTraits = [],
  dropdownStyleReset,
}: RunOptions): SyncRun {
  return SyncCoordinator.init().run({
    isSending,
    intendedState: { alterations },
    previousIdNameMap: previousItemMap(),
    forcedTraits,
    frame,
    dropdownStyleReset,
  });
}

const layoutIdPrefixes = { layoutLeft: "lyl", layoutRight: "lyr" } as const;

// The Layout tab's two managed Tables as the previous ID-name map records them.
export function previousLayoutMap(): IdNameMapGeneric {
  const tables = Object.fromEntries(
    Object.entries(layoutIdPrefixes).map(([table, idPrefix]) => [
      table,
      {
        name: table,
        sheetGid: layoutGid,
        columns: {
          entry: dimensionIds.col(idPrefix, "entry"),
          amount: dimensionIds.col(idPrefix, "amount"),
        },
      },
    ]),
  );
  return {
    tables,
    tabs: {
      [String(layoutGid)]: {
        name: "layout",
        title: "Layout",
        tableIds: managedLayoutTableNames.map(layoutTableId),
      },
    },
  };
}

export function runLayout(alterations: PendingAlterations): SyncRun {
  return SyncCoordinator.init().run({
    isSending: true,
    intendedState: { alterations },
    previousIdNameMap: previousLayoutMap(),
    forcedTraits: [],
    frame,
  });
}

export const noteColIndex = Object.keys(itemColumns).indexOf("optionalNote");

// The note sits below the top row, which sync's head-row read already covers.
export const bodyRowsWithNote: FakeBodyRow<string>[] = [
  { id: itemRowId },
  { id: "r:itm:second", optionalNote: "kept" },
];

export interface ItemAtFlush {
  // Each Table column a read covered on its own; the head-row reads span every column.
  colIndexesRead: number[];
  note: unknown;
}

export interface ItemAtFlushProps {
  isSending: boolean;
  isHostWithholding?: boolean;
  alterations: PendingAlterations;
  columns?: ItemSheetOptions["columns"];
}

// What the run had read of Item, and held of its note, as each flush went out.
export function itemAtEachFlush({
  alterations,
  columns,
  ...props
}: ItemAtFlushProps): ItemAtFlush[] {
  const running: { coordinator?: SyncCoordinator } = {};
  const atFlush: ItemAtFlush[] = [];
  const service = stubSheetsService({
    sheets: [itemSheet({ columns, bodyRows: bodyRowsWithNote })],
    isDryRun: !props.isSending,
    beforeEachBatchUpdate: () => {
      atFlush.push({
        colIndexesRead: itemColIndexesRead(service),
        note: Val.assert(running.coordinator, "the running coordinator")
          .ss.raw.table("item")
          .row(1)
          .cell(noteColIndex)
          .valueOrEmpty(),
      });
    },
  });
  running.coordinator = SyncCoordinator.init();
  running.coordinator.run({
    ...props,
    intendedState: { alterations },
    previousIdNameMap: previousItemMap(),
    forcedTraits: [],
    frame,
  });
  return atFlush;
}

function itemColIndexesRead(service: FakeSheetsService): number[] {
  const colIndexes = new Set<number>();
  for (const resource of service.getByDataFilterCalls) {
    const { dataFilters = [] } = resource as {
      dataFilters?: { gridRange?: ReadGridRange }[];
    };
    for (const { gridRange } of dataFilters) {
      if (gridRange?.sheetId !== itemGid) continue;
      const { startColumnIndex, endColumnIndex } = gridRange;
      if (startColumnIndex === undefined) continue;
      if (endColumnIndex !== startColumnIndex + 1) continue;
      colIndexes.add(startColumnIndex - expectedSheetLayout.startTableColIndex);
    }
  }
  return [...colIndexes].sort((a, b) => a - b);
}

// A send with empty intended state, so each run is the regen alone.
export function runWithNoAlterations(
  coordinator = SyncCoordinator.init(),
): SyncRun {
  return coordinator.run({
    isSending: true,
    intendedState: { alterations: {} },
    previousIdNameMap: { tables: {}, tabs: {} },
    forcedTraits: [],
    frame,
  });
}

export function sendWithNoAlterations(
  coordinator = SyncCoordinator.init(),
): ConfigRegeneration {
  return Val.assert(
    runWithNoAlterations(coordinator).files,
    "files of a send with no alterations",
  );
}

// The emitted line of one Item column, so a test reads the trait the configs carry.
export function emittedItemColumn(
  files: ConfigRegeneration,
  columnName: string,
): string {
  const line = files.columnConfigs
    .split("\n")
    .find((sourceLine) => sourceLine.startsWith(`    "${columnName}": `));
  return Val.assert(line, `emitted ${columnName} column`);
}

export function emptyValueAllowedOf(
  files: ConfigRegeneration,
  columnName: string,
): unknown {
  const match = emittedItemColumn(files, columnName).match(
    /"emptyValueAllowed": (true|false)/,
  );
  return JSON.parse(Val.assert(match?.[1], "emitted emptyValueAllowed"));
}

export function retiredPrefixesOf(files: ConfigRegeneration): unknown {
  const match = files.retiredIdPrefixes.match(/makeRetiredIdPrefixes\((.*)\);/);
  return JSON.parse(Val.assert(match?.[1], "emitted retired prefixes"));
}

export const optionalNoteAlteration: PendingAlterations = {
  tables: {
    item: { columns: { optionalNote: { emptyValueAllowed: false } } },
  },
};
