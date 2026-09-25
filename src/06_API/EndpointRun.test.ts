import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { getColumnTraitByName } from "../01_SpreadsheetSchema/columnConfigsTypes";
import { getSheetTraitByName } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import { SpreadsheetBaseNamed } from "../04_SpreadsheetNamed/ClassBases/SpreadsheetBaseNamed";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  buildGridRows,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import { EndpointRun } from "./EndpointRun";
import type { ActionReturn, Endpoint } from "./Endpoints";

type BatchUpdateCall =
  GoogleAppsScript.Sheets.Schema.BatchUpdateSpreadsheetRequest;

const runItemGid = getSheetTraitByName("runItem", "sheetGid");
const columnIds = (["id", "selected", "startTime", "runStatus"] as const).map(
  (columnName) => getColumnTraitByName("runItem", columnName, "columnId"),
);
const headers = ["ID", "Selected", "Start time", "Run status"];
const selectorColIndex = 1;
const timeLastRanColIndex = 2;
const runStatusColIndex = 3;
const actionRowIndex = 2;
const topDataRowIndex = 4;
const endRowIndex = 9;

const lightYellow = { red: 1, green: 0.949, blue: 0.8 };
const lightGreen = { red: 0.851, green: 0.918, blue: 0.827 };
const lightOrange = { red: 0.99, green: 0.85, blue: 0.7 };
const lightRed = { red: 0.957, green: 0.8, blue: 0.8 };

const timestamp = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/;

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

function runEndpoint(endpoint: Endpoint<"runItem">, isChecked = true) {
  const run = new EndpointRun({
    ...SpreadsheetBaseNamed.initSpreadsheetNamedProps(),
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

function allRequests(calls: BatchUpdateCall[]) {
  return calls.flatMap((call) => call.requests ?? []);
}

function fillRequestsFor(calls: BatchUpdateCall[], colIndex: number) {
  return allRequests(calls).filter(
    (request) => request.repeatCell?.range?.startColumnIndex === colIndex,
  );
}

function fillsFor(calls: BatchUpdateCall[], colIndex: number) {
  return fillRequestsFor(calls, colIndex).map((request) => ({
    startRowIndex: request.repeatCell?.range?.startRowIndex,
    endRowIndex: request.repeatCell?.range?.endRowIndex,
    value: request.repeatCell?.cell?.userEnteredValue?.stringValue,
    backgroundColor:
      request.repeatCell?.cell?.userEnteredFormat?.backgroundColor,
  }));
}

function checkboxFillsFor(calls: BatchUpdateCall[], colIndex: number) {
  return fillRequestsFor(calls, colIndex).map((request) => ({
    startRowIndex: request.repeatCell?.range?.startRowIndex,
    endRowIndex: request.repeatCell?.range?.endRowIndex,
    value: request.repeatCell?.cell?.userEnteredValue?.boolValue,
  }));
}

function cellWritesFor(calls: BatchUpdateCall[], colIndex: number) {
  return allRequests(calls)
    .filter(
      (request) =>
        request.updateCells?.range?.startColumnIndex === colIndex &&
        (request.updateCells.range.startRowIndex ?? 0) >= topDataRowIndex,
    )
    .map((request) => {
      const cell = request.updateCells?.rows?.[0]?.values?.[0];
      return {
        rowIndex: request.updateCells?.range?.startRowIndex,
        value: cell?.userEnteredValue?.stringValue,
        backgroundColor: cell?.userEnteredFormat?.backgroundColor,
      };
    });
}

function actionRowWrites(calls: BatchUpdateCall[]) {
  return allRequests(calls).filter((request) => {
    const range = request.repeatCell?.range ?? request.updateCells?.range;
    return range?.startRowIndex === actionRowIndex;
  });
}

function touchedRowIndexes(calls: BatchUpdateCall[]): number[] {
  const rowIndexes = allRequests(calls)
    .flatMap((request) => {
      const range = request.repeatCell?.range ?? request.updateCells?.range;
      const start = range?.startRowIndex ?? 0;
      const end = range?.endRowIndex ?? start + 1;
      return Array.from({ length: end - start }, (_, i) => start + i);
    })
    .filter((rowIndex) => rowIndex >= topDataRowIndex);
  return [...new Set(rowIndexes)].sort((a, b) => a - b);
}

beforeEach(() => {
  stubLogger();
});

describe("EndpointRun.run, an endpoint with a selector", () => {
  it("stamps the run status into the selected rows only", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(selectiveEndpoint(noOp));

    expect(fillsFor(batchUpdateCalls, runStatusColIndex)).toEqual([
      {
        startRowIndex: 4,
        endRowIndex: 5,
        value: "Running…",
        backgroundColor: lightYellow,
      },
      {
        startRowIndex: 6,
        endRowIndex: 7,
        value: "Running…",
        backgroundColor: lightYellow,
      },
      {
        startRowIndex: 4,
        endRowIndex: 5,
        value: "Succeeded",
        backgroundColor: lightGreen,
      },
      {
        startRowIndex: 6,
        endRowIndex: 7,
        value: "Succeeded",
        backgroundColor: lightGreen,
      },
    ]);
  });

  it("leaves every unselected row completely untouched", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(selectiveEndpoint(noOp));

    expect(touchedRowIndexes(batchUpdateCalls)).toEqual([4, 6]);
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
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(selectiveEndpoint(noOp));
    const writes = fillsFor(batchUpdateCalls, timeLastRanColIndex);

    expect(writes[0]?.value).toMatch(timestamp);
    expect(writes[0]?.backgroundColor).toEqual(lightYellow);
    expect(writes[1]?.value).toMatch(timestamp);
    expect(writes[1]?.backgroundColor).toEqual(lightYellow);
    expect(writes.slice(2)).toEqual([
      {
        startRowIndex: 4,
        endRowIndex: 5,
        value: undefined,
        backgroundColor: lightGreen,
      },
      {
        startRowIndex: 6,
        endRowIndex: 7,
        value: undefined,
        backgroundColor: lightGreen,
      },
    ]);
  });

  it("reads the selection without a round trip of its own", () => {
    const { getByDataFilterCalls } = stubRunItemSheet();

    runEndpoint(selectiveEndpoint(noOp));

    expect(getByDataFilterCalls).toHaveLength(2);
  });
});

describe("EndpointRun.run, the selection a successful run consumes", () => {
  it("unticks the selected rows and no other row", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(selectiveEndpoint(noOp));

    expect(checkboxFillsFor(batchUpdateCalls, selectorColIndex)).toEqual([
      { startRowIndex: 4, endRowIndex: 5, value: false },
      { startRowIndex: 6, endRowIndex: 7, value: false },
    ]);
  });

  it("leaves the ticks alone when the action throws", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(
      selectiveEndpoint(() => {
        throw new Error("no good");
      }),
    );

    expect(checkboxFillsFor(batchUpdateCalls, selectorColIndex)).toEqual([]);
    expect(fillsFor(batchUpdateCalls, runStatusColIndex).at(-1)?.value).toBe(
      "Error: no good",
    );
  });

  it("leaves the ticks alone when the endpoint retains its selection", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(retainingEndpoint(noOp));

    expect(checkboxFillsFor(batchUpdateCalls, selectorColIndex)).toEqual([]);
    expect(fillsFor(batchUpdateCalls, runStatusColIndex).at(-1)?.value).toBe(
      "Succeeded",
    );
  });

  it("unticks them on an untick run too, leaving the entry checkbox alone", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint({ ...selectiveEndpoint(noOp), runOnUncheck: true }, false);

    expect(checkboxFillsFor(batchUpdateCalls, selectorColIndex)).toEqual([
      { startRowIndex: 4, endRowIndex: 5, value: false },
      { startRowIndex: 6, endRowIndex: 7, value: false },
    ]);
    expect(actionRowWrites(batchUpdateCalls)).toEqual([]);
  });
});

describe("EndpointRun.run, an endpoint with no selector", () => {
  it("stamps every table data row through a single fill per state", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(reportingEndpoint(noOp));
    const writes = fillsFor(batchUpdateCalls, timeLastRanColIndex);

    expect(writes[0]?.startRowIndex).toBe(topDataRowIndex);
    expect(writes[0]?.endRowIndex).toBe(endRowIndex);
    expect(writes[0]?.value).toMatch(timestamp);
    expect(writes[0]?.backgroundColor).toEqual(lightYellow);
  });

  it("writes the start time once and only recolours it afterwards", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(reportingEndpoint(noOp));
    const writes = fillsFor(batchUpdateCalls, timeLastRanColIndex);

    expect(writes[1]).toEqual({
      startRowIndex: topDataRowIndex,
      endRowIndex,
      value: undefined,
      backgroundColor: lightGreen,
    });
    expect(writes).toHaveLength(2);
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

  it("still stamps its status across the blank row, so an emptied sheet reports somewhere", () => {
    const { batchUpdateCalls } = stubRunItemSheetWithBlankRow();

    runEndpoint(reportingEndpoint(noOp));

    expect(fillsFor(batchUpdateCalls, runStatusColIndex)[0]).toEqual({
      startRowIndex: topDataRowIndex,
      endRowIndex,
      value: "Running…",
      backgroundColor: lightYellow,
    });
  });
});

describe("EndpointRun.run, the run status message", () => {
  it("writes the action's returned string", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(reportingEndpoint(() => "Built 5 items"));

    expect(fillsFor(batchUpdateCalls, runStatusColIndex).at(-1)?.value).toBe(
      "Built 5 items",
    );
  });

  it("writes Succeeded when the action returns nothing", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(reportingEndpoint(noOp));

    expect(fillsFor(batchUpdateCalls, runStatusColIndex).at(-1)?.value).toBe(
      "Succeeded",
    );
  });
});

describe("EndpointRun.run, the two flushes", () => {
  it("puts the running state on the sheet before the work begins", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(reportingEndpoint(noOp));

    expect(
      fillsFor(batchUpdateCalls.slice(0, 1), runStatusColIndex).map(
        (write) => write.value,
      ),
    ).toEqual(["Running…"]);
    expect(batchUpdateCalls).toHaveLength(2);
  });
});

describe("EndpointRun.run, the start time", () => {
  afterEach(() => {
    vi.useRealTimers();
  });

  it("is the wall-clock time in the spreadsheet's own zone", () => {
    vi.useFakeTimers({ toFake: ["Date"] });
    vi.setSystemTime(new Date("2024-03-15T02:30:00Z"));
    const { batchUpdateCalls } = stubRunItemSheet([4, 6], "Asia/Tokyo");

    runEndpoint(selectiveEndpoint(noOp));

    expect(fillsFor(batchUpdateCalls, timeLastRanColIndex)[0]?.value).toBe(
      "2024-03-15 11:30:00",
    );
  });
});

describe("EndpointRun.run, an endpoint declaring no feedback columns", () => {
  it("emits no stamp at all", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint({
      action: noOp,
      selector: { column: "selected" },
    });

    expect(fillsFor(batchUpdateCalls, timeLastRanColIndex)).toEqual([]);
    expect(fillsFor(batchUpdateCalls, runStatusColIndex)).toEqual([]);
  });
});

describe("EndpointRun.run, a run that fails", () => {
  // Reading a row past the table's last one is a real read on real state.
  function failingAction(ss: Parameters<Endpoint<"runItem">["action"]>[0]) {
    ss.sheet("runItem").row(4).cell("id").updateValue("r:rit:written");
    ss.sheet("runItem").row(endRowIndex).value("id");
  }

  it("writes the error text and red to the selected rows only", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(selectiveEndpoint(failingAction));

    expect(fillsFor(batchUpdateCalls, runStatusColIndex).slice(2)).toEqual([
      {
        startRowIndex: 4,
        endRowIndex: 5,
        value: expect.stringMatching(/^Error: /) as string,
        backgroundColor: lightRed,
      },
      {
        startRowIndex: 6,
        endRowIndex: 7,
        value: expect.stringMatching(/^Error: /) as string,
        backgroundColor: lightRed,
      },
    ]);
    expect(fillsFor(batchUpdateCalls, timeLastRanColIndex).slice(2)).toEqual([
      {
        startRowIndex: 4,
        endRowIndex: 5,
        value: undefined,
        backgroundColor: lightRed,
      },
      {
        startRowIndex: 6,
        endRowIndex: 7,
        value: undefined,
        backgroundColor: lightRed,
      },
    ]);
  });

  it("discards what the action queued before it threw", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(selectiveEndpoint(failingAction));

    const idWrites = batchUpdateCalls
      .flatMap((call) => call.requests ?? [])
      .filter((request) => request.updateCells?.range?.startColumnIndex === 0);
    expect(idWrites).toEqual([]);
  });
});

describe("EndpointRun.run, an empty selection", () => {
  it("runs no action and stamps nothing", () => {
    const { batchUpdateCalls } = stubRunItemSheet([]);
    const calls: string[] = [];

    runEndpoint(
      selectiveEndpoint(() => {
        calls.push("ran");
      }),
    );

    expect(calls).toEqual([]);
    expect(fillsFor(batchUpdateCalls, runStatusColIndex)).toEqual([]);
    expect(batchUpdateCalls).toHaveLength(1);
  });
});

describe("EndpointRun.run, a selector that requires one row", () => {
  it("refuses a selection of two, naming the sheet and how many were ticked", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(oneRowEndpoint(noOp));

    expect(fillsFor(batchUpdateCalls, runStatusColIndex).at(-1)?.value).toBe(
      'Error: This endpoint runs on one row of "Run item" at a time, but 2 are selected.',
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
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(oneRowEndpoint(noOp));

    expect(checkboxFillsFor(batchUpdateCalls, selectorColIndex)).toEqual([]);
  });

  it("reports the refusal as a failed run on every ticked row", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(oneRowEndpoint(noOp));

    expect(fillsFor(batchUpdateCalls, timeLastRanColIndex).slice(2)).toEqual([
      {
        startRowIndex: 4,
        endRowIndex: 5,
        value: undefined,
        backgroundColor: lightRed,
      },
      {
        startRowIndex: 6,
        endRowIndex: 7,
        value: undefined,
        backgroundColor: lightRed,
      },
    ]);
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
    const { batchUpdateCalls } = stubRunItemSheet([6]);

    runEndpoint(oneRowEndpoint(noOp));

    expect(fillsFor(batchUpdateCalls, runStatusColIndex).at(-1)?.value).toBe(
      "Succeeded",
    );
  });
});

describe("EndpointRun.run, a run report naming a state", () => {
  it("writes the warning's own message, since warning has no useful default", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(
      reportingEndpoint(() => ({
        runState: "warning" as const,
        message: "Added 3 of 5",
      })),
    );

    expect(fillsFor(batchUpdateCalls, runStatusColIndex).at(-1)?.value).toBe(
      "Added 3 of 5",
    );
  });

  it("colours both feedback columns, so a sheet with one of them still shows it", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(
      reportingEndpoint(() => ({
        runState: "warning" as const,
        message: "Added 3 of 5",
      })),
    );

    expect(
      fillsFor(batchUpdateCalls, runStatusColIndex).at(-1)?.backgroundColor,
    ).toEqual(lightOrange);
    expect(
      fillsFor(batchUpdateCalls, timeLastRanColIndex).at(-1)?.backgroundColor,
    ).toEqual(lightOrange);
  });

  it("leaves the start time written once, recolouring it without a value", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(
      reportingEndpoint(() => ({
        runState: "warning" as const,
        message: "Added 3 of 5",
      })),
    );
    const writes = fillsFor(batchUpdateCalls, timeLastRanColIndex);

    expect(writes[0]?.value).toMatch(timestamp);
    expect(writes[1]?.value).toBeUndefined();
    expect(writes).toHaveLength(2);
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

  it("writes each named row's own message and colour into its own cell", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(reportingEndpoint(twoRowsFailed));

    expect(cellWritesFor(batchUpdateCalls, runStatusColIndex)).toEqual([
      { rowIndex: 5, value: "No such row", backgroundColor: lightRed },
      { rowIndex: 7, value: "Amount is blank", backgroundColor: lightRed },
    ]);
  });

  it("colours the named rows' start-time cells without rewriting the time", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(reportingEndpoint(twoRowsFailed));

    expect(cellWritesFor(batchUpdateCalls, timeLastRanColIndex)).toEqual([
      { rowIndex: 5, value: undefined, backgroundColor: lightRed },
      { rowIndex: 7, value: undefined, backgroundColor: lightRed },
    ]);
  });

  it("defaults the unnamed rows to success when no state is named", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(reportingEndpoint(twoRowsFailed));

    expect(fillsFor(batchUpdateCalls, runStatusColIndex).at(-1)).toEqual({
      startRowIndex: topDataRowIndex,
      endRowIndex,
      value: "Succeeded",
      backgroundColor: lightGreen,
    });
  });

  it("gives the unnamed rows the state named beside the map", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(
      reportingEndpoint(() => ({
        runState: "warning" as const,
        message: "Added 3 of 5",
        ...twoRowsFailed(),
      })),
    );

    expect(fillsFor(batchUpdateCalls, runStatusColIndex).at(-1)).toEqual({
      startRowIndex: topDataRowIndex,
      endRowIndex,
      value: "Added 3 of 5",
      backgroundColor: lightOrange,
    });
  });

  it("lets a named row be warned rather than failed", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(
      reportingEndpoint(() => ({
        rows: new Map([
          [5, { runState: "warning" as const, message: "Check this one" }],
        ]),
      })),
    );

    expect(cellWritesFor(batchUpdateCalls, runStatusColIndex)).toEqual([
      { rowIndex: 5, value: "Check this one", backgroundColor: lightOrange },
    ]);
  });

  it("keeps the rest of the run's writes, since a returned failure is not a throw", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(
      reportingEndpoint((ss) => {
        ss.sheet("runItem").row(4).cell("id").updateValue("r:rit:written");
        return twoRowsFailed();
      }),
    );

    expect(cellWritesFor(batchUpdateCalls, 0)).toEqual([
      { rowIndex: 4, value: "r:rit:written", backgroundColor: undefined },
    ]);
  });

  it("lets a selector endpoint flag some of its selected rows and not others", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(
      selectiveEndpoint(() => ({
        rows: new Map([
          [6, { runState: "failure" as const, message: "No such row" }],
        ]),
      })),
    );

    expect(cellWritesFor(batchUpdateCalls, runStatusColIndex)).toEqual([
      { rowIndex: 6, value: "No such row", backgroundColor: lightRed },
    ]);
  });

  it("fails the run when a key is not a data row of the sheet", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(
      reportingEndpoint(() => ({
        rows: new Map([
          [endRowIndex, { runState: "failure" as const, message: "Nowhere" }],
        ]),
      })),
    );

    expect(fillsFor(batchUpdateCalls, runStatusColIndex).at(-1)).toEqual({
      startRowIndex: topDataRowIndex,
      endRowIndex,
      value: `Error: Row ${endRowIndex} is not a data row of "Run item", so this run cannot report into it.`,
      backgroundColor: lightRed,
    });
  });

  it("costs the run no round trip of its own", () => {
    const { batchUpdateCalls, getByDataFilterCalls } = stubRunItemSheet();

    runEndpoint(reportingEndpoint(twoRowsFailed));

    expect(getByDataFilterCalls).toHaveLength(1);
    expect(batchUpdateCalls).toHaveLength(2);
  });

  it("leaves a deleted row's delete standing when the same run also names it", () => {
    const { batchUpdateCalls } = stubRunItemSheet();

    runEndpoint(
      reportingEndpoint((ss) => {
        ss.sheet("runItem").row(5).delete();
        return twoRowsFailed();
      }),
    );

    expect(
      allRequests(batchUpdateCalls).filter(
        (request) => request.deleteDimension?.range?.startIndex === 5,
      ),
    ).toHaveLength(1);
    expect(cellWritesFor(batchUpdateCalls, runStatusColIndex)).toEqual([
      { rowIndex: 7, value: "Amount is blank", backgroundColor: lightRed },
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
