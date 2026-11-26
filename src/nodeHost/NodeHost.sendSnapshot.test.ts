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
import { stubSheetsService } from "../testSupport/fakeSheetsService.ts";
import {
  type FakeBodyRow,
  fakeTableSheet,
} from "../testSupport/fakeSheetsService/fakeTableSheet.ts";
import type { HeldList } from "./EnteredDataGate.ts";
import { heldGroups } from "./heldGroups.ts";
import { NodeHost } from "./NodeHost.ts";

const spreadsheetId = "spreadsheet-under-test";
const gateTarget = "app:spreadsheet-under-test";

// The Widget Table, where the layout puts it: column IDs in row 0, header in row 3, data from row 4.
// Gid 0, as the fake's Table range carries no sheetId, which the gate reads as 0.
const widget = { gid: 0, title: "Widget", endRow: 6 };
const widgetColumns = {
  id: { columnId: "col-id", header: "ID" },
  name: { columnId: "col-name", header: "Name" },
  amount: { columnId: "col-amount", header: "Amount" },
  note: { columnId: "col-note", header: "Note" },
};
type WidgetColumn = keyof typeof widgetColumns;
const widgetColumnNames: WidgetColumn[] = ["id", "name", "amount", "note"];
const widgetBodyRows: FakeBodyRow<WidgetColumn>[] = [
  { id: "row-1", name: "Sprocket", amount: 1200 },
  { id: "row-2", name: "Gizmo", amount: 50, note: "late" },
];
// A tab with no Table the run reads.
const scratch = { gid: 6, title: "Scratch" };

function widgetCol(column: WidgetColumn): number {
  return widgetColumnNames.indexOf(column);
}

function seedFake() {
  return stubSheetsService({
    sheets: [
      fakeTableSheet.build({
        sheetId: widget.gid,
        title: widget.title,
        name: widget.title,
        columnConfigs: widgetColumns,
        columnNames: widgetColumnNames,
        bodyRows: widgetBodyRows,
      }),
      { sheetId: scratch.gid, title: scratch.title, rows: [["kept"]] },
    ],
  });
}

interface SnapshotCase {
  flushes: GoogleRequest[][];
  forcedTraits?: ForcedTrait[];
}

interface HostSeed {
  isDryRun: boolean;
  confirmedList: HeldList | undefined;
  forcedTraits: ForcedTrait[] | undefined;
}

function seedHost({
  isDryRun,
  confirmedList,
  forcedTraits,
}: HostSeed): NodeHost {
  return NodeHost.init({
    configs: installedConfigs(),
    spreadsheetId,
    transport: fakeSheetsHttpTransport(seedFake()),
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
}

function readWidget(api: GoogleSheetsAPI): void {
  api.fetchGrid(
    [
      {
        sheetId: widget.gid,
        startRowIndex: SheetIndex.row(0),
        endRowIndex: SheetIndex.row(widget.endRow),
        startColumnIndex: SheetIndex.col(0),
        endColumnIndex: SheetIndex.col(widgetColumnNames.length),
      },
    ],
    { includeProgrammaticFacts: false },
  );
}

function runFlushes(host: NodeHost, flushes: GoogleRequest[][]): void {
  const api = host.googleSheetsAPI;
  readWidget(api);
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
  const host = seedHost({
    isDryRun: true,
    confirmedList: undefined,
    forcedTraits,
  });
  runFlushes(host, flushes);
  return host;
}

// A dry run for the token, then the send it covers.
function confirmedSend(snapshotCase: SnapshotCase) {
  const heldList = Val.assert(dryRun(snapshotCase).heldList, "heldList");
  const host = seedHost({
    isDryRun: false,
    confirmedList: heldList,
    forcedTraits: snapshotCase.forcedTraits,
  });
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

const overwriteSprocket = writeCell(widget.gid, 4, widgetCol("amount"), {
  numberValue: 900,
});
const noteSprocket = writeCell(widget.gid, 4, widgetCol("note"), {
  stringValue: "memo",
});
const deleteNoteColumn: GoogleRequest = {
  deleteRange: {
    range: {
      sheetId: widget.gid,
      startColumnIndex: widgetCol("note"),
      endColumnIndex: widgetCol("note") + 1,
    },
    shiftDimension: "COLUMNS",
  },
};
const replaceSprocket: GoogleRequest = {
  findReplace: { find: "Sprocket", replacement: "Cog", allSheets: true },
};
const overwriteScratch = writeCell(scratch.gid, 0, 0, {
  stringValue: "gone",
});

// The Widget tab as the fixture seeds it, blank cells as batchGet gives them.
const blankRow = widgetColumnNames.map(() => "");
const seededRows: unknown[][] = [
  widgetColumnNames.map((column) => widgetColumns[column].columnId),
  blankRow,
  blankRow,
  widgetColumnNames.map((column) => widgetColumns[column].header),
  ...widgetBodyRows.map((row) =>
    widgetColumnNames.map((column) => row[column] ?? ""),
  ),
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
      flushes: [[overwriteSprocket, noteSprocket]],
    });

    expect(snapshot?.columns).toEqual([
      {
        sheetGid: widget.gid,
        tabTitle: widget.title,
        table: widget.title,
        header: "Amount",
        columnId: "col-amount",
        colIndex: widgetCol("amount"),
        values: seededColumn(widgetCol("amount")),
      },
    ]);
    const fingerprintCols = [
      widgetCol("id"),
      widgetCol("name"),
      widgetCol("note"),
    ];
    const before = seededFingerprints(fingerprintCols);
    expect(snapshot?.fingerprints).toEqual([
      {
        sheetGid: widget.gid,
        tabTitle: widget.title,
        colIndexes: fingerprintCols,
        before,
        after: before.with(4, ["row-1", "Sprocket", "memo"]),
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
        colIndex: widgetCol("note"),
        values: seededColumn(widgetCol("note")),
      }),
    ]);
    expect(snapshot?.fingerprints[0]?.colIndexes).toEqual([
      widgetCol("id"),
      widgetCol("name"),
      widgetCol("amount"),
    ]);
  });

  it("holds each column as it stood before the first flush that changes it", () => {
    const { snapshot } = confirmedSend({
      flushes: [
        [overwriteSprocket],
        [writeCell(widget.gid, 5, widgetCol("amount"), { numberValue: 1 })],
      ],
    });

    expect(snapshot?.columns.map(({ values }) => values)).toEqual([
      seededColumn(widgetCol("amount")),
    ]);
  });

  it("names a findReplace and a cell no Table read covers as undone only by version history", () => {
    const { snapshot } = confirmedSend({
      flushes: [[replaceSprocket], [overwriteScratch]],
    });

    expect(snapshot?.columns).toEqual([]);
    expect(snapshot?.versionHistoryItems.map(({ kind }) => kind)).toEqual([
      "findReplace",
      "unreadCell",
    ]);
  });

  it("names the same items in the dry run's held groups as undone only by version history", () => {
    const heldList = dryRun({
      flushes: [[overwriteSprocket], [replaceSprocket], [overwriteScratch]],
    }).heldList;

    const groups = heldGroups(Val.assert(heldList, "heldList").flushes);

    expect(groups).toEqual([
      expect.objectContaining({
        subject: { table: widget.title, column: "Amount" },
      }),
      expect.objectContaining({
        subject: { tab: `gid ${scratch.gid}`, path: "A" },
        isUndoneOnlyByVersionHistory: true,
      }),
      expect.objectContaining({
        what: expect.stringContaining("Sprocket"),
        isUndoneOnlyByVersionHistory: true,
      }),
    ]);
    expect(groups[0]).not.toHaveProperty("isUndoneOnlyByVersionHistory");
  });

  it("snapshots a forced send's traits though the send holds no cell", () => {
    const trait = { table: widget.title, column: "Amount", trait: "valueName" };

    const { snapshot } = confirmedSend({
      flushes: [[noteSprocket]],
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
      dryRun({ flushes: [[overwriteSprocket]] }).sendSnapshot(),
    ).toBeUndefined();
  });

  it("takes none on a send with nothing held", () => {
    expect(
      confirmedSend({ flushes: [[noteSprocket]] }).snapshot,
    ).toBeUndefined();
  });
});
