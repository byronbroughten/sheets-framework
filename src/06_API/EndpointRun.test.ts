import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { getColumnTraitByName } from "../01_SpreadsheetSchema/columnConfigsTypes";
import { getSheetTraitByName } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import { SpreadsheetBaseNamed } from "../04_SpreadsheetNamed/ClassBases/SpreadsheetBaseNamed";
import type { SpreadsheetNamed } from "../04_SpreadsheetNamed/SpreadsheetNamed";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  buildGridRows,
  type FakeCell,
  type FakeCellValue,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import type { FakeGridView } from "../testSupport/fakeSheetsService/gridView";
import { EndpointRun } from "./EndpointRun";
import type { ActionReturn, Endpoint, EndpointsAll } from "./Endpoints";
import { feedbackColumnIdsOf } from "./feedbackColumnIds";

type Color = GoogleAppsScript.Sheets.Schema.Color;

const runItemGid = getSheetTraitByName("runItem", "sheetGid");
const columnIds = (["id", "selected", "startTime", "runStatus"] as const).map(
  (columnName) => getColumnTraitByName("runItem", columnName, "columnId"),
);
const headers = ["ID", "Selected", "Start time", "Run status"];
const idColIndex = 0;
const selectorColIndex = 1;
const timeLastRanColIndex = 2;
const runStatusColIndex = 3;
const resultColIndex = 4;
const actionRowIndex = 2;
const topDataRowIndex = 4;
const endRowIndex = 9;

const lightYellow = { red: 1, green: 0.949, blue: 0.8 };
const lightGreen = { red: 0.851, green: 0.918, blue: 0.827 };
const lightOrange = { red: 0.99, green: 0.85, blue: 0.7 };
const lightRed = { red: 0.957, green: 0.8, blue: 0.8 };

const timestamp = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/;
const runStartUtc = new Date("2024-03-15T02:30:00Z");
const runStartInSheetZone = "2024-03-14 21:30:00";

const actionRowWithEntryTicked = [null, null, true, null];
const selectionAsTicked = [true, false, true, false, false];

// Rows 4 and 6 are ticked; 5, 7 and 8 are the rows a selective run must not touch.
function stubRunItemSheet(
  checkedRowIndexes: number[] = [4, 6],
  timeZone?: string,
) {
  const dataRow = (rowIndex: number) => [
    `r:rit:row${rowIndex}`,
    checkedRowIndexes.includes(rowIndex),
    "",
    "",
  ];
  return stubSheetsService({
    timeZone,
    sheets: [
      {
        sheetId: runItemGid,
        title: "Run item",
        rows: buildGridRows({
          0: columnIds,
          [actionRowIndex]: actionRowWithEntryTicked,
          3: headers,
          4: dataRow(4),
          5: dataRow(5),
          6: dataRow(6),
          7: dataRow(7),
          8: dataRow(8),
        }),
        table: { endRowIndex },
      },
    ],
  });
}

// Row 6 is the blank row an emptied-then-refilled sheet would be left with.
function stubRunItemSheetWithBlankRow() {
  const dataRow = (rowIndex: number) => [`r:rit:row${rowIndex}`, false, "", ""];
  return stubSheetsService({
    sheets: [
      {
        sheetId: runItemGid,
        title: "Run item",
        rows: buildGridRows({
          0: columnIds,
          [actionRowIndex]: actionRowWithEntryTicked,
          3: headers,
          4: dataRow(4),
          5: dataRow(5),
          6: [null, null, null, null],
          7: dataRow(7),
          8: dataRow(8),
        }),
        table: { endRowIndex },
      },
    ],
  });
}

// One data row, nothing in it but what `topRow` holds: blank, as an earlier run emptied it, unless told otherwise.
function stubOneRowRunItemSheet(topRow: Partial<Record<string, FakeCell>>) {
  const columnNames = [
    "id",
    "selected",
    "startTime",
    "runStatus",
    "result",
  ] as const;
  return stubSheetsService({
    sheets: [
      {
        sheetId: runItemGid,
        title: "Run item",
        rows: buildGridRows({
          0: columnNames.map((columnName) =>
            getColumnTraitByName("runItem", columnName, "columnId"),
          ),
          [actionRowIndex]: actionRowWithEntryTicked,
          3: ["ID", "Selected", "Start time", "Run status", "Result"],
          [topDataRowIndex]: columnNames.map(
            (columnName) => topRow[columnName] ?? null,
          ),
        }),
        table: { endRowIndex: topDataRowIndex + 1 },
      },
    ],
  });
}

interface RunOptions {
  isChecked?: boolean;
  alsoDeclared?: EndpointsAll;
}

function runEndpoint(
  endpoint: Endpoint<"runItem">,
  { isChecked = true, alsoDeclared = {} }: RunOptions = {},
) {
  const run = new EndpointRun({
    ...SpreadsheetBaseNamed.initSpreadsheetNamedProps(
      feedbackColumnIdsOf({ ...alsoDeclared, runItem_startTime: endpoint }),
    ),
    sheetName: "runItem",
    entryColumnName: "startTime",
    endpoint,
  });
  run.sheet.identified.meta.ensureColumnIdsAreFetched();
  run.run(isChecked);
}

function reportingEndpoint(
  action: Endpoint<"runItem">["action"],
): Endpoint<"runItem"> {
  return {
    action,
    timeLastRan: "startTime",
    runStatus: "runStatus",
  };
}

function selectiveEndpoint(
  action: Endpoint<"runItem">["action"],
): Endpoint<"runItem"> {
  return {
    ...reportingEndpoint(action),
    selector: { column: "selected" },
  };
}

function oneRowEndpoint(
  action: Endpoint<"runItem">["action"],
): Endpoint<"runItem"> {
  return {
    ...reportingEndpoint(action),
    selector: { column: "selected", requireOneRow: true },
  };
}

function retainingEndpoint(
  action: Endpoint<"runItem">["action"],
): Endpoint<"runItem"> {
  return {
    ...reportingEndpoint(action),
    selector: { column: "selected", retainSelection: true },
  };
}

function noOp() {}

function appendRow(ss: SpreadsheetNamed): void {
  const sheet = ss.sheet("runItem");
  sheet.row(topDataRowIndex).prepFetchFull();
  ss.fetchAllPrepped();
  sheet.appendRowWithVals({ result: "appended" });
}

// Every row down to the grid's end, so a row appended beneath the blank one shows.
function idColumnValues(grid: FakeGridView): FakeCellValue[] {
  return grid
    .sheet(runItemGid)
    .values({
      startRowIndex: topDataRowIndex,
      startColumnIndex: idColIndex,
      endColumnIndex: idColIndex + 1,
    })
    .flat();
}

function startClockAtRunStart() {
  vi.useFakeTimers({ toFake: ["Date"] });
  vi.setSystemTime(runStartUtc);
}

// The clock moves on while the action works, so a rewritten start time would show.
function withClockMovingDuringAction(
  action: Endpoint<"runItem">["action"],
): Endpoint<"runItem">["action"] {
  return (ss, props) => {
    vi.setSystemTime(new Date(runStartUtc.getTime() + 15 * 60 * 1000));
    return action(ss, props);
  };
}

function columnCells(grid: FakeGridView, colIndex: number): FakeCell[] {
  return grid
    .sheet(runItemGid)
    .rows({
      startRowIndex: topDataRowIndex,
      endRowIndex,
      startColumnIndex: colIndex,
      endColumnIndex: colIndex + 1,
    })
    .flat();
}

function dataRows(grid: FakeGridView): FakeCell[][] {
  return grid
    .sheet(runItemGid)
    .rows({ startRowIndex: topDataRowIndex, endRowIndex });
}

function entryCell(grid: FakeGridView): FakeCell {
  return grid.sheet(runItemGid).cell(actionRowIndex, timeLastRanColIndex);
}

function stamp(value: string, backgroundColor: Color): FakeCell {
  return { value, backgroundColor };
}

function stampedTime(backgroundColor: Color): FakeCell {
  return { value: expect.stringMatching(timestamp) as string, backgroundColor };
}

// Rows 4 and 6 hold the cell; the unselected rows keep the fixture's blank.
function onSelectedRows(cell: FakeCell): FakeCell[] {
  return [cell, "", cell, "", ""];
}

function unselectedRows(rows: FakeCell[][]): (FakeCell[] | undefined)[] {
  return [5, 7, 8].map((rowIndex) => rows[rowIndex - topDataRowIndex]);
}

function allRows(cell: FakeCell): FakeCell[] {
  return Array.from({ length: endRowIndex - topDataRowIndex }, () => cell);
}

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.useRealTimers();
});

describe("EndpointRun.run, an endpoint with a selector", () => {
  it("shows the running state on the selected rows only while the action works", () => {
    const { grid } = stubRunItemSheet();
    let statusesWhileRunning: FakeCell[] = [];

    runEndpoint(
      selectiveEndpoint(() => {
        statusesWhileRunning = columnCells(grid, runStatusColIndex);
      }),
    );

    expect(statusesWhileRunning).toEqual(
      onSelectedRows(stamp("Running…", lightYellow)),
    );
  });

  it("leaves the run status on the selected rows only", () => {
    const { grid } = stubRunItemSheet();

    runEndpoint(selectiveEndpoint(noOp));

    expect(columnCells(grid, runStatusColIndex)).toEqual(
      onSelectedRows(stamp("Succeeded", lightGreen)),
    );
  });

  it("leaves every unselected row completely untouched", () => {
    const { grid } = stubRunItemSheet();
    const before = dataRows(grid);

    runEndpoint(selectiveEndpoint(noOp));

    expect(unselectedRows(dataRows(grid))).toEqual(unselectedRows(before));
  });

  it("hands the action exactly the rows it stamps", () => {
    stubRunItemSheet();
    let received: number[] = [];

    runEndpoint(
      selectiveEndpoint((_ss, { selectedRowIndexes }) => {
        received = selectedRowIndexes;
      }),
    );

    expect(received).toEqual([4, 6]);
  });

  it("shows the run state on every selected row's start-time cell", () => {
    const { grid } = stubRunItemSheet();
    let startTimesWhileRunning: FakeCell[] = [];

    runEndpoint(
      selectiveEndpoint(() => {
        startTimesWhileRunning = columnCells(grid, timeLastRanColIndex);
      }),
    );

    expect(startTimesWhileRunning).toEqual(
      onSelectedRows(stampedTime(lightYellow)),
    );
    expect(columnCells(grid, timeLastRanColIndex)).toEqual(
      onSelectedRows(stampedTime(lightGreen)),
    );
  });

  it("reads the selection without a round trip of its own", () => {
    const { getByDataFilterCalls } = stubRunItemSheet();

    runEndpoint(selectiveEndpoint(noOp));

    expect(getByDataFilterCalls).toHaveLength(2);
  });
});

describe("EndpointRun.run, the selection a successful run consumes", () => {
  it("unticks the selected rows", () => {
    const { grid } = stubRunItemSheet();

    runEndpoint(selectiveEndpoint(noOp));

    expect(columnCells(grid, selectorColIndex)).toEqual(allRows(false));
  });

  it("leaves the ticks alone when the action throws", () => {
    const { grid } = stubRunItemSheet();

    runEndpoint(
      selectiveEndpoint(() => {
        throw new Error("no good");
      }),
    );

    expect(columnCells(grid, selectorColIndex)).toEqual(selectionAsTicked);
    expect(columnCells(grid, runStatusColIndex)[0]).toEqual(
      stamp("Error: no good", lightRed),
    );
  });

  it("leaves the ticks alone when the endpoint retains its selection", () => {
    const { grid } = stubRunItemSheet();

    runEndpoint(retainingEndpoint(noOp));

    expect(columnCells(grid, selectorColIndex)).toEqual(selectionAsTicked);
    expect(columnCells(grid, runStatusColIndex)[0]).toEqual(
      stamp("Succeeded", lightGreen),
    );
  });

  it("unticks them on an untick run too, leaving the entry checkbox alone", () => {
    const { grid } = stubRunItemSheet();

    runEndpoint(
      { ...selectiveEndpoint(noOp), runOnUncheck: true },
      { isChecked: false },
    );

    expect(columnCells(grid, selectorColIndex)).toEqual(allRows(false));
    expect(entryCell(grid)).toBe(true);
  });

  it("unticks a selected row the action ticked again", () => {
    const { grid } = stubRunItemSheet();

    runEndpoint(
      selectiveEndpoint((ss) => {
        ss.sheet("runItem").row(4).updateValue("selected", true);
      }),
    );

    expect(columnCells(grid, selectorColIndex)).toEqual(allRows(false));
  });

  it("clears a button's entry checkbox for the next click", () => {
    const { grid } = stubRunItemSheet();

    runEndpoint(selectiveEndpoint(noOp));

    expect(entryCell(grid)).toBe(false);
  });
});

describe("EndpointRun.run, an endpoint with no selector", () => {
  it("stamps the start time on every table data row while the action works", () => {
    const { grid } = stubRunItemSheet();
    let startTimesWhileRunning: FakeCell[] = [];

    runEndpoint(
      reportingEndpoint(() => {
        startTimesWhileRunning = columnCells(grid, timeLastRanColIndex);
      }),
    );

    expect(startTimesWhileRunning).toEqual(allRows(stampedTime(lightYellow)));
  });

  it("keeps the start time it wrote, only recolouring it afterwards", () => {
    startClockAtRunStart();
    const { grid } = stubRunItemSheet();

    runEndpoint(reportingEndpoint(withClockMovingDuringAction(noOp)));

    expect(columnCells(grid, timeLastRanColIndex)).toEqual(
      allRows(stamp(runStartInSheetZone, lightGreen)),
    );
  });

  it("tells the action about every data row", () => {
    stubRunItemSheet();
    let received: number[] = [];

    runEndpoint(
      reportingEndpoint((_ss, { selectedRowIndexes }) => {
        received = selectedRowIndexes;
      }),
    );

    expect(received).toEqual([4, 5, 6, 7, 8]);
  });

  it("costs no read of its own, since nothing is prepped", () => {
    const { getByDataFilterCalls } = stubRunItemSheet();

    runEndpoint(reportingEndpoint(noOp));

    expect(getByDataFilterCalls).toHaveLength(1);
  });

  it("keeps a blank row out of the rows it hands the action", () => {
    stubRunItemSheetWithBlankRow();
    let received: number[] = [];

    runEndpoint(
      reportingEndpoint((_ss, { selectedRowIndexes }) => {
        received = selectedRowIndexes;
      }),
    );

    expect(received).toEqual([4, 5, 7, 8]);
  });

  it("shows its outcome on the row left by an action that empties the sheet, and reuses that row", () => {
    const { grid } = stubOneRowRunItemSheet({
      id: "r:rit:row4",
      result: "old",
    });

    runEndpoint(
      reportingEndpoint((ss) => {
        ss.sheet("runItem").DELETE_ALL_DATA_ROWS();
      }),
    );
    expect(
      grid.sheet(runItemGid).cell(topDataRowIndex, runStatusColIndex),
    ).toEqual(stamp("Succeeded", lightGreen));
    runEndpoint(reportingEndpoint(appendRow));

    expect(idColumnValues(grid)).toEqual([expect.stringMatching(/^r:rit:/)]);
  });

  it("still stamps its status across the blank row, which stays reusable, so an emptied sheet reports somewhere", () => {
    const { grid } = stubOneRowRunItemSheet({});
    let statusesWhileRunning: FakeCell[] = [];

    runEndpoint(
      reportingEndpoint(() => {
        statusesWhileRunning = grid
          .sheet(runItemGid)
          .rows({
            startRowIndex: topDataRowIndex,
            startColumnIndex: runStatusColIndex,
            endColumnIndex: runStatusColIndex + 1,
          })
          .flat();
      }),
    );
    expect(statusesWhileRunning).toEqual([stamp("Running…", lightYellow)]);
    expect(
      grid.sheet(runItemGid).cell(topDataRowIndex, runStatusColIndex),
    ).toEqual(stamp("Succeeded", lightGreen));
    runEndpoint(reportingEndpoint(appendRow));

    expect(idColumnValues(grid)).toEqual([expect.stringMatching(/^r:rit:/)]);
  });
});

describe("EndpointRun.run, a blank row another endpoint stamped", () => {
  const stampingEndpoint = {
    action: noOp,
    runStatus: "result",
  } as const satisfies Endpoint<"runItem">;

  it("is reused by this endpoint's append, since every declared endpoint's feedback columns are skipped", () => {
    const { grid } = stubOneRowRunItemSheet({ result: "Succeeded" });

    runEndpoint(reportingEndpoint(appendRow), {
      alsoDeclared: { runItem_selected: stampingEndpoint },
    });

    expect(idColumnValues(grid)).toEqual([expect.stringMatching(/^r:rit:/)]);
  });

  it("is not reused when no declared endpoint reports into that column", () => {
    const { grid } = stubOneRowRunItemSheet({ result: "Succeeded" });

    runEndpoint(reportingEndpoint(appendRow));

    expect(
      grid.sheet(runItemGid).values({
        startRowIndex: topDataRowIndex,
        endRowIndex: topDataRowIndex + 2,
        startColumnIndex: idColIndex,
        endColumnIndex: resultColIndex + 1,
      }),
    ).toEqual([
      [null, null, expect.anything(), expect.anything(), "Succeeded"],
      [
        expect.stringMatching(/^r:rit:/),
        false,
        expect.anything(),
        expect.anything(),
        "appended",
      ],
    ]);
  });
});

describe("EndpointRun.run, the run status message", () => {
  it("writes the action's returned string", () => {
    const { grid } = stubRunItemSheet();

    runEndpoint(reportingEndpoint(() => "Built 5 items"));

    expect(columnCells(grid, runStatusColIndex)).toEqual(
      allRows(stamp("Built 5 items", lightGreen)),
    );
  });

  it("writes Succeeded when the action returns nothing", () => {
    const { grid } = stubRunItemSheet();

    runEndpoint(reportingEndpoint(noOp));

    expect(columnCells(grid, runStatusColIndex)).toEqual(
      allRows(stamp("Succeeded", lightGreen)),
    );
  });
});

describe("EndpointRun.run, the two flushes", () => {
  it("puts the running state on the sheet before the work begins", () => {
    const { grid } = stubRunItemSheet();
    let statusesWhileRunning: FakeCell[] = [];

    runEndpoint(
      reportingEndpoint(() => {
        statusesWhileRunning = columnCells(grid, runStatusColIndex);
      }),
    );

    expect(statusesWhileRunning).toEqual(
      allRows(stamp("Running…", lightYellow)),
    );
  });

  it("costs two batch updates, one before the work and one after", () => {
    const { batchUpdateCount } = stubRunItemSheet();

    runEndpoint(reportingEndpoint(noOp));

    expect(batchUpdateCount()).toBe(2);
  });
});

describe("EndpointRun.run, the start time", () => {
  it("is the wall-clock time in the spreadsheet's own zone", () => {
    startClockAtRunStart();
    const { grid } = stubRunItemSheet([4, 6], "Asia/Tokyo");

    runEndpoint(selectiveEndpoint(noOp));

    expect(columnCells(grid, timeLastRanColIndex)[0]).toEqual(
      stamp("2024-03-15 11:30:00", lightGreen),
    );
  });
});

describe("EndpointRun.run, an endpoint declaring no feedback columns", () => {
  it("leaves the start-time and status columns as they were", () => {
    const { grid } = stubRunItemSheet();
    const startTimes = columnCells(grid, timeLastRanColIndex);
    const statuses = columnCells(grid, runStatusColIndex);

    runEndpoint({
      action: noOp,
      selector: { column: "selected" },
    });

    expect(columnCells(grid, timeLastRanColIndex)).toEqual(startTimes);
    expect(columnCells(grid, runStatusColIndex)).toEqual(statuses);
  });
});

describe("EndpointRun.run, a run that fails", () => {
  // Reading a row past the table's last one is a real read on real state.
  function failingAction(ss: Parameters<Endpoint<"runItem">["action"]>[0]) {
    ss.sheet("runItem").row(4).cell("id").updateValue("r:rit:written");
    ss.sheet("runItem").row(endRowIndex).value("id");
  }

  it("writes the error text and red to the selected rows only", () => {
    const { grid } = stubRunItemSheet();

    runEndpoint(selectiveEndpoint(failingAction));

    const error = {
      value: expect.stringMatching(/^Error: /) as string,
      backgroundColor: lightRed,
    };
    expect(columnCells(grid, runStatusColIndex)).toEqual(onSelectedRows(error));
    expect(columnCells(grid, timeLastRanColIndex)).toEqual(
      onSelectedRows(stampedTime(lightRed)),
    );
  });

  it("discards what the action queued before it threw", () => {
    const { grid } = stubRunItemSheet();

    runEndpoint(selectiveEndpoint(failingAction));

    expect(columnCells(grid, idColIndex)[0]).toBe("r:rit:row4");
  });
});

describe("EndpointRun.run, an empty selection", () => {
  it("runs no action and stamps nothing", () => {
    const { grid } = stubRunItemSheet([]);
    const before = dataRows(grid);
    const calls: string[] = [];

    runEndpoint(
      selectiveEndpoint(() => {
        calls.push("ran");
      }),
    );

    expect(calls).toEqual([]);
    expect(dataRows(grid)).toEqual(before);
  });
});

describe("EndpointRun.run, a selector that requires one row", () => {
  it("refuses a selection of two, naming the sheet and how many were ticked", () => {
    const { grid } = stubRunItemSheet();

    runEndpoint(oneRowEndpoint(noOp));

    expect(columnCells(grid, runStatusColIndex)[0]).toEqual(
      stamp(
        'Error: This endpoint runs on one row of "Run item" at a time, but 2 are selected.',
        lightRed,
      ),
    );
  });

  it("never reaches the action", () => {
    stubRunItemSheet();
    const calls: string[] = [];

    runEndpoint(
      oneRowEndpoint(() => {
        calls.push("ran");
      }),
    );

    expect(calls).toEqual([]);
  });

  it("leaves the ticks alone, so the extras can be unticked and the run retried", () => {
    const { grid } = stubRunItemSheet();

    runEndpoint(oneRowEndpoint(noOp));

    expect(columnCells(grid, selectorColIndex)).toEqual(selectionAsTicked);
  });

  it("reports the refusal as a failed run on every ticked row", () => {
    const { grid } = stubRunItemSheet();

    runEndpoint(oneRowEndpoint(noOp));

    expect(columnCells(grid, timeLastRanColIndex)).toEqual(
      onSelectedRows(stampedTime(lightRed)),
    );
  });

  it("hands the action its one row when exactly one is ticked", () => {
    stubRunItemSheet([6]);
    let received: number[] = [];

    runEndpoint(
      oneRowEndpoint((_ss, { selectedRowIndexes }) => {
        received = selectedRowIndexes;
      }),
    );

    expect(received).toEqual([6]);
  });

  it("succeeds on that one row", () => {
    const { grid } = stubRunItemSheet([6]);

    runEndpoint(oneRowEndpoint(noOp));

    expect(columnCells(grid, runStatusColIndex)).toEqual([
      "",
      "",
      stamp("Succeeded", lightGreen),
      "",
      "",
    ]);
  });
});

describe("EndpointRun.run, a run report naming a state", () => {
  function warned() {
    return { runState: "warning" as const, message: "Added 3 of 5" };
  }

  it("writes the warning's own message, since warning has no useful default", () => {
    const { grid } = stubRunItemSheet();

    runEndpoint(reportingEndpoint(warned));

    expect(columnCells(grid, runStatusColIndex)).toEqual(
      allRows(stamp("Added 3 of 5", lightOrange)),
    );
  });

  it("colours both feedback columns, so a sheet with one of them still shows it", () => {
    const { grid } = stubRunItemSheet();

    runEndpoint(reportingEndpoint(warned));

    expect(columnCells(grid, runStatusColIndex)).toEqual(
      allRows(stamp("Added 3 of 5", lightOrange)),
    );
    expect(columnCells(grid, timeLastRanColIndex)).toEqual(
      allRows(stampedTime(lightOrange)),
    );
  });

  it("keeps the start time it wrote, recolouring it without a value", () => {
    startClockAtRunStart();
    const { grid } = stubRunItemSheet();

    runEndpoint(reportingEndpoint(withClockMovingDuringAction(warned)));

    expect(columnCells(grid, timeLastRanColIndex)).toEqual(
      allRows(stamp(runStartInSheetZone, lightOrange)),
    );
  });
});

describe("EndpointRun.run, a run report naming rows", () => {
  function twoRowsFailed() {
    return {
      rows: new Map([
        [5, { runState: "failure" as const, message: "No such row" }],
        [7, { runState: "failure" as const, message: "Amount is blank" }],
      ]),
    };
  }

  it("overwrites what the action wrote into the run-status column, keeping the named rows' own reports", () => {
    const { grid } = stubRunItemSheet();

    runEndpoint(
      reportingEndpoint((ss) => {
        ss.sheet("runItem").row(4).updateValue("runStatus", "mine");
        ss.sheet("runItem").row(5).updateValue("runStatus", "mine");
        return twoRowsFailed();
      }),
    );

    expect(columnCells(grid, runStatusColIndex)).toEqual([
      stamp("Succeeded", lightGreen),
      stamp("No such row", lightRed),
      stamp("Succeeded", lightGreen),
      stamp("Amount is blank", lightRed),
      stamp("Succeeded", lightGreen),
    ]);
  });

  it("writes each named row's own message and colour into its own cell", () => {
    const { grid } = stubRunItemSheet();

    runEndpoint(reportingEndpoint(twoRowsFailed));

    expect(columnCells(grid, runStatusColIndex)).toEqual([
      stamp("Succeeded", lightGreen),
      stamp("No such row", lightRed),
      stamp("Succeeded", lightGreen),
      stamp("Amount is blank", lightRed),
      stamp("Succeeded", lightGreen),
    ]);
  });

  it("defaults the unnamed rows to success when no state is named", () => {
    const { grid } = stubRunItemSheet();

    runEndpoint(reportingEndpoint(twoRowsFailed));

    const statuses = columnCells(grid, runStatusColIndex);
    expect([statuses[0], statuses[2], statuses[4]]).toEqual(
      Array(3).fill(stamp("Succeeded", lightGreen)),
    );
  });

  it("colours the named rows' start-time cells without rewriting the time", () => {
    startClockAtRunStart();
    const { grid } = stubRunItemSheet();

    runEndpoint(reportingEndpoint(withClockMovingDuringAction(twoRowsFailed)));

    expect(columnCells(grid, timeLastRanColIndex)).toEqual([
      stamp(runStartInSheetZone, lightGreen),
      stamp(runStartInSheetZone, lightRed),
      stamp(runStartInSheetZone, lightGreen),
      stamp(runStartInSheetZone, lightRed),
      stamp(runStartInSheetZone, lightGreen),
    ]);
  });

  it("gives the unnamed rows the state named beside the map", () => {
    const { grid } = stubRunItemSheet();

    runEndpoint(
      reportingEndpoint(() => ({
        runState: "warning" as const,
        message: "Added 3 of 5",
        ...twoRowsFailed(),
      })),
    );

    expect(columnCells(grid, runStatusColIndex)).toEqual([
      stamp("Added 3 of 5", lightOrange),
      stamp("No such row", lightRed),
      stamp("Added 3 of 5", lightOrange),
      stamp("Amount is blank", lightRed),
      stamp("Added 3 of 5", lightOrange),
    ]);
  });

  it("lets a named row be warned rather than failed", () => {
    const { grid } = stubRunItemSheet();

    runEndpoint(
      reportingEndpoint(() => ({
        rows: new Map([
          [5, { runState: "warning" as const, message: "Check this one" }],
        ]),
      })),
    );

    expect(columnCells(grid, runStatusColIndex)[1]).toEqual(
      stamp("Check this one", lightOrange),
    );
  });

  it("keeps the rest of the run's writes, since a returned failure is not a throw", () => {
    const { grid } = stubRunItemSheet();

    runEndpoint(
      reportingEndpoint((ss) => {
        ss.sheet("runItem").row(4).cell("id").updateValue("r:rit:written");
        return twoRowsFailed();
      }),
    );

    expect(columnCells(grid, idColIndex)[0]).toBe("r:rit:written");
  });

  it("lets a selector endpoint flag some of its selected rows and not others", () => {
    const { grid } = stubRunItemSheet();

    runEndpoint(
      selectiveEndpoint(() => ({
        rows: new Map([
          [6, { runState: "failure" as const, message: "No such row" }],
        ]),
      })),
    );

    expect(columnCells(grid, runStatusColIndex)).toEqual([
      stamp("Succeeded", lightGreen),
      "",
      stamp("No such row", lightRed),
      "",
      "",
    ]);
  });

  it("fails the run when a key is not a data row of the sheet", () => {
    const { grid } = stubRunItemSheet();

    runEndpoint(
      reportingEndpoint(() => ({
        rows: new Map([
          [endRowIndex, { runState: "failure" as const, message: "Nowhere" }],
        ]),
      })),
    );

    expect(columnCells(grid, runStatusColIndex)).toEqual(
      allRows(
        stamp(
          `Error: Row ${endRowIndex} is not a data row of "Run item", so this run cannot report into it.`,
          lightRed,
        ),
      ),
    );
  });

  it("costs the run no read and no batch update of its own", () => {
    const { batchUpdateCount, getByDataFilterCalls } = stubRunItemSheet();

    runEndpoint(reportingEndpoint(twoRowsFailed));

    expect(getByDataFilterCalls).toHaveLength(1);
    expect(batchUpdateCount()).toBe(2);
  });

  it("leaves a deleted row deleted when the same run also names it", () => {
    const { grid } = stubRunItemSheet();

    runEndpoint(
      reportingEndpoint((ss) => {
        ss.sheet("runItem").row(5).delete();
        return twoRowsFailed();
      }),
    );

    expect(dataRows(grid).map(([id]) => id)).toEqual([
      "r:rit:row4",
      "r:rit:row6",
      "r:rit:row7",
      "r:rit:row8",
      null,
    ]);
    expect(columnCells(grid, runStatusColIndex)).toEqual([
      stamp("Succeeded", lightGreen),
      stamp("Succeeded", lightGreen),
      stamp("Amount is blank", lightRed),
      stamp("Succeeded", lightGreen),
      null,
    ]);
  });
});

describe("the run report's type", () => {
  it("refuses a warning or a failure that carries no message", () => {
    // @ts-expect-error warning has no fallback sentence, so the type demands one
    const warned: ActionReturn = { runState: "warning" };
    // @ts-expect-error and failure has none either
    const failed: ActionReturn = { runState: "failure" };

    expect([warned, failed]).toEqual([
      { runState: "warning" },
      { runState: "failure" },
    ]);
  });

  it("takes a success with no message, and a message with no state", () => {
    const bare: ActionReturn = { runState: "success" };
    const messaged: ActionReturn = { message: "Built 5 items" };

    expect([bare, messaged]).toEqual([
      { runState: "success" },
      { message: "Built 5 items" },
    ]);
  });
});
