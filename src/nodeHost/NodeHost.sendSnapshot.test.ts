import { Val } from "@byronbroughten/utils/val";
import { describe, expect, it, vi } from "vitest";

import {
  googleRawRequest,
  type GoogleRequest,
  type GoogleSheetsAPI,
} from "../00_Source/GoogleSheets/GoogleSheetsAPI.ts";
import { SheetIndex } from "../00_Source/RawSource/SheetIndex.ts";
import { installedConfigs } from "../01_SpreadsheetSchema/configReaders/configRegister.ts";
import type { ForcedTrait } from "../01_SpreadsheetSchema/configReaders/syncRunResult.ts";
import { fakeSheetsHttpTransport } from "../testSupport/fakeSheetsHttpTransport.ts";
import {
  buildGridRows,
  type FakeSheetsService,
  stubSheetsService,
} from "../testSupport/fakeSheetsService.ts";
import type { HeldList } from "./EnteredDataGate.ts";
import { heldGroups } from "./heldGroups.ts";
import { NodeHost } from "./NodeHost.ts";

const spreadsheetId = "spreadsheet-under-test";
const gateTarget = "app:spreadsheet-under-test";

// The Ledger Table: column IDs in row 0, header in row 3 (ID, Name, Amount, Note), data from row 4.
// Gid 0, as the fake's Table range carries no sheetId, which the gate reads as 0.
const ledger = { gid: 0, title: "Ledger", headerRow: 3, endRow: 6 };
const ledgerCol = { id: 0, name: 1, amount: 2, note: 3 };
const ledgerRows = buildGridRows({
  3: ["ID", "Name", "Amount", "Note"],
  4: ["row-1", "Rent", 1200],
  5: ["row-2", "Fee", 50, "late"],
});
const ledgerColumnIds = ["col-id", "col-name", "col-amount", "col-note"];
// A tab with no Table the run reads.
const scratch = { gid: 6, title: "Scratch" };

function seedFake(): FakeSheetsService {
  return stubSheetsService({
    sheets: [
      {
        sheetId: ledger.gid,
        title: ledger.title,
        rows: ledgerRows,
        table: {
          name: "Ledger",
          startRowIndex: ledger.headerRow,
          startColumnIndex: 0,
          endRowIndex: ledger.endRow,
          headRows: { 3: ledgerColumnIds },
        },
      },
      { sheetId: scratch.gid, title: scratch.title, rows: [["kept"]] },
    ],
  });
}

interface SnapshotCase {
  flushes: GoogleRequest[][];
  forcedTraits?: ForcedTrait[];
}

function seedHost(
  isDryRun: boolean,
  confirmedList: HeldList | undefined,
  forcedTraits: ForcedTrait[] | undefined,
) {
  const fake = seedFake();
  const host = NodeHost.init({
    configs: installedConfigs(),
    spreadsheetId,
    transport: fakeSheetsHttpTransport(fake),
    isDryRun,
    log: vi.fn(),
    enteredDataGate: {
      target: gateTarget,
      confirmedList,
      dropdownStyleReset: undefined,
      forcedTraits,
      hash: (input) => `hash(${input})`,
    },
  });
  return { host, fake, api: host.googleSheetsAPI };
}

function readLedger(api: GoogleSheetsAPI): void {
  api.fetchGrid(
    [
      {
        sheetId: ledger.gid,
        startRowIndex: SheetIndex.row(0),
        endRowIndex: SheetIndex.row(ledger.endRow),
        startColumnIndex: SheetIndex.col(0),
        endColumnIndex: SheetIndex.col(4),
      },
    ],
    { includeProgrammaticFacts: false },
  );
}

function runFlushes(host: NodeHost, flushes: GoogleRequest[][]): void {
  const api = host.googleSheetsAPI;
  readLedger(api);
  for (const requests of flushes) {
    api.flush(
      requests.map((request) => ({
        kind: "raw",
        request: googleRawRequest(request),
      })),
    );
  }
}

function dryRun({ flushes, forcedTraits }: SnapshotCase) {
  const { host } = seedHost(true, undefined, forcedTraits);
  runFlushes(host, flushes);
  return host;
}

// A dry run for the token, then the send it covers.
function confirmedSend(snapshotCase: SnapshotCase) {
  const heldList = Val.assert(dryRun(snapshotCase).heldList, "heldList");
  const { host } = seedHost(false, heldList, snapshotCase.forcedTraits);
  runFlushes(host, snapshotCase.flushes);
  return { snapshot: host.sendSnapshot() };
}

function writeCell(
  sheetId: number,
  rowIndex: number,
  colIndex: number,
  value: GoogleAppsScript.Sheets.Schema.ExtendedValue,
): GoogleRequest {
  return {
    updateCells: {
      range: {
        sheetId,
        startRowIndex: rowIndex,
        endRowIndex: rowIndex + 1,
        startColumnIndex: colIndex,
        endColumnIndex: colIndex + 1,
      },
      rows: [{ values: [{ userEnteredValue: value }] }],
      fields: "userEnteredValue",
    },
  };
}

const overwriteRent = writeCell(ledger.gid, 4, ledgerCol.amount, {
  numberValue: 900,
});
const noteRent = writeCell(ledger.gid, 4, ledgerCol.note, {
  stringValue: "memo",
});
const deleteNoteColumn: GoogleRequest = {
  deleteRange: {
    range: {
      sheetId: ledger.gid,
      startColumnIndex: ledgerCol.note,
      endColumnIndex: ledgerCol.note + 1,
    },
    shiftDimension: "COLUMNS",
  },
};
const replaceRent: GoogleRequest = {
  findReplace: { find: "Rent", replacement: "Lease", allSheets: true },
};
const overwriteScratch = writeCell(scratch.gid, 0, 0, {
  stringValue: "gone",
});

const seededRows: unknown[][] = [
  ledgerColumnIds,
  ["", "", "", ""],
  ["", "", "", ""],
  ["ID", "Name", "Amount", "Note"],
  ["row-1", "Rent", 1200, ""],
  ["row-2", "Fee", 50, "late"],
];

function seededColumn(colIndex: number): unknown[] {
  return seededRows.map((row) => row[colIndex]);
}

function seededFingerprints(colIndexes: number[]): unknown[][] {
  return seededRows.map((row) => colIndexes.map((index) => row[index]));
}

describe("NodeHost.sendSnapshot", () => {
  it("holds an overwritten column whole by its column ID, and fingerprints the tab before and after the send", () => {
    const { snapshot } = confirmedSend({
      flushes: [[overwriteRent, noteRent]],
    });

    expect(snapshot?.columns).toEqual([
      {
        sheetGid: ledger.gid,
        tabTitle: ledger.title,
        table: "Ledger",
        header: "Amount",
        columnId: "col-amount",
        colIndex: ledgerCol.amount,
        values: seededColumn(ledgerCol.amount),
      },
    ]);
    const fingerprintCols = [ledgerCol.id, ledgerCol.name, ledgerCol.note];
    const before = seededFingerprints(fingerprintCols);
    expect(snapshot?.fingerprints).toEqual([
      {
        sheetGid: ledger.gid,
        tabTitle: ledger.title,
        colIndexes: fingerprintCols,
        before,
        after: before.with(4, ["row-1", "Rent", "memo"]),
      },
    ]);
    expect(snapshot?.versionHistoryItems).toEqual([]);
  });

  it("holds a deleted column's values, head rows and position, so it can be recreated", () => {
    const { snapshot } = confirmedSend({ flushes: [[deleteNoteColumn]] });

    expect(snapshot?.columns).toEqual([
      expect.objectContaining({
        header: "Note",
        columnId: "col-note",
        colIndex: ledgerCol.note,
        values: ["col-note", "", "", "Note", "", "late"],
      }),
    ]);
    expect(snapshot?.fingerprints[0]?.colIndexes).toEqual([
      ledgerCol.id,
      ledgerCol.name,
      ledgerCol.amount,
    ]);
  });

  it("holds each column as it stood before the first flush that changes it", () => {
    const { snapshot } = confirmedSend({
      flushes: [
        [overwriteRent],
        [writeCell(ledger.gid, 5, ledgerCol.amount, { numberValue: 1 })],
      ],
    });

    expect(snapshot?.columns.map(({ values }) => values)).toEqual([
      seededColumn(ledgerCol.amount),
    ]);
  });

  it("names a findReplace and a cell no Table read covers as undone only by version history", () => {
    const { snapshot } = confirmedSend({
      flushes: [[replaceRent], [overwriteScratch]],
    });

    expect(snapshot?.columns).toEqual([]);
    expect(snapshot?.versionHistoryItems.map(({ kind }) => kind)).toEqual([
      "findReplace",
      "unreadCell",
    ]);
  });

  it("names the same items in the dry run's held groups as undone only by version history", () => {
    const heldList = dryRun({
      flushes: [[overwriteRent], [replaceRent], [overwriteScratch]],
    }).heldList;

    const groups = heldGroups(Val.assert(heldList, "heldList").flushes);

    expect(groups).toEqual([
      expect.objectContaining({
        subject: { table: "Ledger", column: "Amount" },
      }),
      expect.objectContaining({
        subject: { tab: `gid ${scratch.gid}`, path: "A" },
        isUndoneOnlyByVersionHistory: true,
      }),
      expect.objectContaining({
        what: expect.stringContaining("Rent"),
        isUndoneOnlyByVersionHistory: true,
      }),
    ]);
    expect(groups[0]).not.toHaveProperty("isUndoneOnlyByVersionHistory");
  });

  it("snapshots a forced send's traits though the send holds no cell", () => {
    const trait = { table: "Ledger", column: "Amount", trait: "valueName" };

    const { snapshot } = confirmedSend({
      flushes: [[noteRent]],
      forcedTraits: [trait],
    });

    expect(snapshot).toEqual({
      columns: [],
      fingerprints: [],
      versionHistoryItems: [{ kind: "forcedTrait", trait }],
    });
  });

  it("takes none on a dry run", () => {
    expect(
      dryRun({ flushes: [[overwriteRent]] }).sendSnapshot(),
    ).toBeUndefined();
  });

  it("takes none on a send with nothing held", () => {
    expect(confirmedSend({ flushes: [[noteRent]] }).snapshot).toBeUndefined();
  });
});
