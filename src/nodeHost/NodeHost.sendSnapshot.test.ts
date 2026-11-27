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
  type FakeTableColumn,
  fakeTableSheet,
} from "../testSupport/fakeSheetsService/fakeTableSheet.ts";
import type { ConfirmedList, HeldList } from "./EnteredDataGate.ts";
import { GateReport } from "./GateReport.ts";
import { heldGroups } from "./heldGroups.ts";
import { NodeHost } from "./NodeHost.ts";

const spreadsheetId = "spreadsheet-under-test";
const gateTarget = "app:spreadsheet-under-test";

// The Widget Table, where the layout puts it: column IDs in row 0, header in row 3, data from row 4.
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
const versionHistoryLine =
  "undone only by the spreadsheet's version history: the send snapshot can't hold it";
// A tab with no Table the run reads.
const scratch = { gid: 6, title: "Scratch" };

function widgetCol(column: WidgetColumn): number {
  return widgetColumnNames.indexOf(column);
}

function seedFake(blankColumnIds: WidgetColumn[]) {
  const columnConfigs: Record<WidgetColumn, FakeTableColumn> = {
    ...widgetColumns,
  };
  for (const column of blankColumnIds) {
    columnConfigs[column] = { ...widgetColumns[column], columnId: "" };
  }
  return stubSheetsService({
    sheets: [
      fakeTableSheet.build({
        sheetId: widget.gid,
        title: widget.title,
        name: widget.title,
        columnConfigs,
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
  // No read before the flushes, so every cell they write is unread.
  isWidgetUnread?: boolean;
  // The read starts at the header row, so it never sees the column IDs.
  isColumnIdRowUnread?: boolean;
  // Columns whose column-ID cell is blank on the sheet.
  blankColumnIds?: WidgetColumn[];
}

interface HostSeed {
  isDryRun: boolean;
  confirmedList: ConfirmedList | undefined;
  snapshotCase: SnapshotCase;
}

function seedHost({
  isDryRun,
  confirmedList,
  snapshotCase,
}: HostSeed): NodeHost {
  return NodeHost.init({
    configs: installedConfigs(),
    spreadsheetId,
    transport: fakeSheetsHttpTransport(
      seedFake(snapshotCase.blankColumnIds ?? []),
    ),
    isDryRun,
    log: vi.fn(),
    enteredDataGate: {
      target: gateTarget,
      confirmedList,
      dropdownStyleReset: undefined,
      forcedTraits: snapshotCase.forcedTraits,
      hash: (input) => `hash(${input})`,
    },
  });
}

function readWidget(api: GoogleSheetsAPI, isColumnIdRowUnread = false): void {
  api.fetchGrid(
    [
      {
        sheetId: widget.gid,
        startRowIndex: SheetIndex.row(isColumnIdRowUnread ? 3 : 0),
        endRowIndex: SheetIndex.row(widget.endRow),
        startColumnIndex: SheetIndex.col(0),
        endColumnIndex: SheetIndex.col(widgetColumnNames.length),
      },
    ],
    { includeProgrammaticFacts: false },
  );
}

function runFlushes(
  host: NodeHost,
  { flushes, isWidgetUnread, isColumnIdRowUnread }: SnapshotCase,
): void {
  const api = host.googleSheetsAPI;
  if (!isWidgetUnread) readWidget(api, isColumnIdRowUnread);
  for (const requests of flushes) {
    api.flush(
      requests.map((request) => ({
        kind: "raw",
        request: googleRawRequest(request),
      })),
    );
  }
}

function dryRun(snapshotCase: SnapshotCase) {
  const host = seedHost({
    isDryRun: true,
    confirmedList: undefined,
    snapshotCase,
  });
  runFlushes(host, snapshotCase);
  return host;
}

function dryRunHeldList(snapshotCase: SnapshotCase): HeldList {
  return Val.assert(dryRun(snapshotCase).heldList, "heldList");
}

// The held block as the dry run prints it.
function dryRunHeldText(snapshotCase: SnapshotCase): string {
  const heldList = dryRunHeldList(snapshotCase);
  return GateReport.heldSections({
    block: {
      groups: heldGroups(heldList),
      token: heldList.token,
      fullListPath: ".probe/sync-held.json",
    },
    forcedGroups: [],
  }).join("\n");
}

// A dry run for the token, then the send it covers, given the held list as the held file keeps it.
function confirmedSend(snapshotCase: SnapshotCase) {
  const { token, flushes } = dryRunHeldList(snapshotCase);
  const host = seedHost({
    isDryRun: false,
    confirmedList: { token, flushes },
    snapshotCase,
  });
  runFlushes(host, snapshotCase);
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
const deleteNameColumn: GoogleRequest = {
  deleteRange: {
    range: {
      sheetId: widget.gid,
      startColumnIndex: widgetCol("name"),
      endColumnIndex: widgetCol("name") + 1,
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
  it("holds an overwritten column whole by its column ID, with its Table's ID, and fingerprints the tab before and after the send", () => {
    const { snapshot } = confirmedSend({
      flushes: [[overwriteSprocket, noteSprocket]],
    });

    expect(snapshot?.columns).toEqual([
      {
        sheetGid: widget.gid,
        tabTitle: widget.title,
        table: widget.title,
        tableId: `fake-table-${widget.gid}`,
        header: "Amount",
        columnId: "col-amount",
        colIndex: widgetCol("amount"),
        columnIdRowIndex: 0,
        idColIndex: widgetCol("id"),
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

  it("names a later flush's held cell as undone only by version history once an earlier flush shifted columns", () => {
    // Gizmo's note, one column left of where the read saw it.
    const overwriteShiftedNote = writeCell(widget.gid, 5, widgetCol("amount"), {
      stringValue: "early",
    });

    const { snapshot } = confirmedSend({
      flushes: [[deleteNameColumn], [overwriteShiftedNote]],
    });

    expect(snapshot?.columns.map(({ header }) => header)).toEqual(["Name"]);
    expect(snapshot?.versionHistoryItems).toEqual([
      expect.objectContaining({
        kind: "enteredData",
        cell: expect.objectContaining({ column: "Note" }),
      }),
    ]);
  });

  it("names a findReplace and a cell no Table read covers as undone only by version history", () => {
    const { snapshot } = confirmedSend({
      flushes: [[replaceSprocket], [overwriteScratch]],
      isWidgetUnread: true,
    });

    expect(snapshot?.columns).toEqual([]);
    expect(snapshot?.versionHistoryItems.map(({ kind }) => kind)).toEqual([
      "findReplace",
      "unreadCell",
    ]);
  });

  it("names the same items in the dry run's held groups as undone only by version history", () => {
    const groups = heldGroups(
      dryRunHeldList({ flushes: [[overwriteSprocket], [replaceSprocket]] }),
    );

    expect(groups).toEqual([
      expect.objectContaining({
        subject: { table: widget.title, column: "Amount" },
      }),
      expect.objectContaining({
        what: expect.stringContaining("Sprocket"),
        isUndoneOnlyByVersionHistory: true,
      }),
    ]);
    expect(groups[0]).not.toHaveProperty("isUndoneOnlyByVersionHistory");
  });

  it("names in the dry run a held column whose column ID is blank and a held cell in a flush after a column shift, as the snapshot leaves them to version history", () => {
    // Gizmo's note, one column left of where the read saw it.
    const overwriteShiftedNote = writeCell(widget.gid, 5, widgetCol("amount"), {
      stringValue: "early",
    });
    const snapshotCase = {
      flushes: [[overwriteSprocket, deleteNameColumn], [overwriteShiftedNote]],
      blankColumnIds: ["amount"],
    } satisfies SnapshotCase;

    const text = dryRunHeldText(snapshotCase);

    expect(text).toContain(
      [
        "Widget › Amount  the send overwrites or deletes entered data in 1 cell:",
        `  ${versionHistoryLine}`,
      ].join("\n"),
    );
    expect(text).toContain(
      [
        "Widget › Note  the send overwrites or deletes entered data in 1 cell:",
        `  ${versionHistoryLine}`,
      ].join("\n"),
    );
    expect(text).toContain(
      [
        "Widget › Name  the send overwrites or deletes entered data in 2 cells:",
        "  row-1  ",
      ].join("\n"),
    );
    const { snapshot } = confirmedSend(snapshotCase);
    expect(snapshot?.columns.map(({ header }) => header)).toEqual(["Name"]);
    expect(
      snapshot?.versionHistoryItems.map((item) =>
        "cell" in item ? item.cell.column : item.kind,
      ),
    ).toEqual(["Amount", "Note"]);
  });

  it("names in the dry run a held cell whose column can't be resolved, as no read covered its tab, as the snapshot leaves it to version history", () => {
    const snapshotCase = {
      flushes: [[overwriteScratch]],
      isWidgetUnread: true,
    } satisfies SnapshotCase;

    const text = dryRunHeldText(snapshotCase);

    expect(text).toContain(
      [
        `tab gid ${scratch.gid} › A  the send writes 1 cell no read covered, so they may hold entered data:`,
        `  ${versionHistoryLine}`,
      ].join("\n"),
    );
    const { snapshot } = confirmedSend(snapshotCase);
    expect(snapshot?.columns).toEqual([]);
    expect(snapshot?.versionHistoryItems.map(({ kind }) => kind)).toEqual([
      "unreadCell",
    ]);
  });

  it("keeps a column whose column-ID cell the run never read, and the send leaves it to version history when that cell is blank", () => {
    const snapshotCase = {
      flushes: [[overwriteSprocket]],
      blankColumnIds: ["amount"],
      isColumnIdRowUnread: true,
    } satisfies SnapshotCase;

    const text = dryRunHeldText(snapshotCase);

    expect(text).toContain(
      "Widget › Amount  the send overwrites or deletes entered data in 1 cell:",
    );
    expect(text).not.toContain(versionHistoryLine);
    const { snapshot } = confirmedSend(snapshotCase);
    expect(snapshot?.columns).toEqual([]);
    expect(snapshot?.versionHistoryItems).toEqual([
      expect.objectContaining({
        kind: "enteredData",
        cell: expect.objectContaining({ column: "Amount" }),
      }),
    ]);
  });

  it("prints a column as two groups when the snapshot holds one flush's cells and leaves another's to version history", () => {
    // Gizmo's amount, one column left once the Name column is deleted.
    const overwriteShiftedGizmo = writeCell(widget.gid, 5, widgetCol("name"), {
      numberValue: 7,
    });
    const snapshotCase = {
      flushes: [[overwriteSprocket, deleteNameColumn], [overwriteShiftedGizmo]],
    } satisfies SnapshotCase;

    const amountHeadings = dryRunHeldText(snapshotCase)
      .split("\n")
      .flatMap((line, index, lines) =>
        line.startsWith("Widget › Amount") ? [[line, lines[index + 1]]] : [],
      );

    expect(amountHeadings).toEqual([
      [
        "Widget › Amount  the send overwrites or deletes entered data in 1 cell:",
        '  row-1  "1200"  → "900"',
      ],
      [
        "Widget › Amount  the send overwrites or deletes entered data in 1 cell:",
        `  ${versionHistoryLine}`,
      ],
    ]);
    const { snapshot } = confirmedSend(snapshotCase);
    expect(snapshot?.columns.map(({ header }) => header)).toEqual([
      "Amount",
      "Name",
    ]);
    expect(snapshot?.versionHistoryItems).toEqual([
      expect.objectContaining({
        cell: expect.objectContaining({ column: "Amount" }),
      }),
    ]);
  });

  it("gives the same token whether or not the snapshot leaves the held cells to version history", () => {
    const held = dryRunHeldList({ flushes: [[overwriteSprocket]] });
    const blankIdHeld = dryRunHeldList({
      flushes: [[overwriteSprocket]],
      blankColumnIds: ["amount"],
    });

    expect(blankIdHeld.flushes).toEqual(held.flushes);
    expect(blankIdHeld.versionHistoryItems).not.toEqual(
      held.versionHistoryItems,
    );
    expect(blankIdHeld.token).toBe(held.token);
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
