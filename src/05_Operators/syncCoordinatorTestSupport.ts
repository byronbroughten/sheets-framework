import { Val } from "@byronbroughten/utils/val";

import { installedConfigs } from "../01_SpreadsheetSchema/configReaders/configRegister";
import { type IdNameMapGeneric } from "../01_SpreadsheetSchema/configReaders/idNameMapTypes";
import { type PendingAlterations } from "../01_SpreadsheetSchema/configReaders/intendedState";
import {
  type ForcedTrait,
  type SyncRunFrame,
} from "../01_SpreadsheetSchema/configReaders/syncRunResult";
import { getTableTraitByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { expectedSheetLayout } from "../testSupport/expectedSheetLayout";
import {
  type FakeSheetProperties,
  type FakeSheetsService,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import {
  type FakeBodyRow,
  type FakeTableColumn,
  fakeTableSheet,
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

interface ItemSheetOptions {
  columns?: Record<string, FakeTableColumn>; // Replaces a configured column by name, or adds one after them.
  startRowIndex?: number;
  columnTypes?: Record<number, string>; // By absolute column index.
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
  });
}

export function seedFixture(
  options: {
    item?: ItemSheetOptions;
    extraSheets?: FakeSheetProperties[];
  } = {},
): FakeSheetsService {
  return stubSheetsService({
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

export interface RunOptions {
  isSending: boolean;
  alterations?: PendingAlterations;
  forcedTraits?: ForcedTrait[];
}

export function runItem({
  isSending,
  alterations = {},
  forcedTraits = [],
}: RunOptions): SyncRun {
  return SyncCoordinator.init().run({
    isSending,
    intendedState: { alterations },
    previousIdNameMap: previousItemMap(),
    forcedTraits,
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

export const optionalNoteAlteration: PendingAlterations = {
  tables: {
    item: { columns: { optionalNote: { emptyValueAllowed: false } } },
  },
};
