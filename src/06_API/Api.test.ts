import { beforeEach, describe, expect, it, vi } from "vitest";

import type { SheetEdit } from "../00_Source/PlatformEvents/sheetEdit";
import { SheetIndex } from "../00_Source/RawSource/SheetIndex";
import {
  getColumnTraitByName,
  getSheetColumnNames,
} from "../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import { installedConfigs } from "../01_SpreadsheetSchema/configReaders/configRegister";
import { getTableTraitByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import {
  layoutGid,
  layoutOrigins,
  layoutSheet,
  layoutTableId,
} from "../02_SpreadsheetRaw/spreadsheetRawTestSupport";
import { expectedSheetLayout } from "../testSupport/expectedSheetLayout";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  buildGridRows,
  type FakeCell,
  stubSheetsService,
} from "../testSupport/fakeSheetsService";
import type { FakeGridView } from "../testSupport/fakeSheetsService/gridView";
import { Api } from "./Api";
import type { Endpoints } from "./Endpoints";

const runItemGid = getTableTraitByName("runItem", "sheetGid");
// The last column is deliberately left without a column id.
const columnIds = [
  getColumnTraitByName("runItem", "id", "columnId"),
  getColumnTraitByName("runItem", "selected", "columnId"),
  getColumnTraitByName("runItem", "result", "columnId"),
  "",
];
const idColIndex = 0;
const twoWayColIndex = 1;
const buttonColIndex = 2;
const blankIdColIndex = 3;
const actionRowIndex = expectedSheetLayout.actionRowIndex;
const endRowIndex = 7;

function stubRunItemSheet(actionRowAsClicked: FakeCell[] = []) {
  return stubSheetsService({
    sheets: [
      {
        sheetId: runItemGid,
        title: "Run item",
        rows: buildGridRows({
          0: columnIds,
          [actionRowIndex]: actionRowAsClicked,
          3: ["ID", "Selected", "Result", ""],
          4: ["r:rit:row4", false, "", ""],
          5: ["r:rit:row5", false, "", ""],
          6: ["r:rit:row6", false, "", ""],
        }),
        table: { endRowIndex },
      },
    ],
  });
}

function tickedAt(colIndex: number): FakeCell[] {
  return columnIds.map((_, index) => (index === colIndex ? true : null));
}

function actionRowEdit(colIndex: number, value: string): SheetEdit {
  return {
    sheetGid: runItemGid,
    rowIndexBase0: expectedSheetLayout.actionRowIndex,
    colIndexBase0: SheetIndex.col(colIndex),
    value,
  };
}

function trackingEndpoints(calls: string[]): Endpoints {
  return {
    runItem_selected: {
      action: (_ss, { isChecked }) => {
        calls.push(`twoWay:${isChecked}`);
      },
      runOnUncheck: true,
    },
    runItem_result: {
      action: () => {
        calls.push("button");
      },
    },
  };
}

function actionRowCells(grid: FakeGridView): FakeCell[] {
  return grid
    .sheet(runItemGid)
    .rows({ startRowIndex: actionRowIndex, endRowIndex: actionRowIndex + 1 })
    .flat();
}

beforeEach(() => {
  stubLogger();
});

const configs = installedConfigs();

describe("Api.handleSheetEdit, the entry call", () => {
  it("supplies the installed configs before anything reads them", async () => {
    vi.resetModules();
    const fresh = await import("./Api");
    const installSource = vi.fn();
    expect(() =>
      fresh.Api.handleSheetEdit(
        { configs, endpoints: {} },
        {
          ...actionRowEdit(twoWayColIndex, "TRUE"),
          rowIndexBase0: SheetIndex.row(endRowIndex),
        },
        installSource,
      ),
    ).not.toThrow();
    expect(installSource).not.toHaveBeenCalled();
  });
  it("installs the source and dispatches a suspected API call", () => {
    stubRunItemSheet();
    const calls: string[] = [];
    const installSource = vi.fn();
    Api.handleSheetEdit(
      { configs, endpoints: trackingEndpoints(calls) },
      actionRowEdit(buttonColIndex, "TRUE"),
      installSource,
    );
    expect(installSource).toHaveBeenCalledOnce();
    expect(calls).toEqual(["button"]);
  });
});

describe("Api.handleSheetChange", () => {
  it("does nothing for a change the platform module doesn't name", () => {
    const installSource = vi.fn();
    expect(
      Api.handleSheetChange(
        { configs, endpoints: {} },
        undefined,
        installSource,
      ),
    ).toBeUndefined();
    expect(installSource).not.toHaveBeenCalled();
  });
  it("installs the source and returns the floor notice for a renamed Value Config", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: getTableTraitByName("valueConfig", "sheetGid"),
          title: "Values",
        },
      ],
    });
    const installSource = vi.fn();
    expect(
      Api.handleSheetChange({ configs, endpoints: {} }, "other", installSource),
    ).toEqual({
      title: "Value Config is managed",
      message:
        'This tab keeps the name "Value Config". Your rename will switch back the next time configs sync.',
      untilClosed: false,
    });
    expect(installSource).toHaveBeenCalledOnce();
  });
  it("returns no floor notice when the floor tabs are intact", () => {
    stubSheetsService({
      sheets: [
        {
          sheetId: getTableTraitByName("valueConfig", "sheetGid"),
          title: "Value Config",
        },
      ],
    });
    expect(
      Api.handleSheetChange({ configs, endpoints: {} }, "other", vi.fn()),
    ).toBeUndefined();
  });
});

describe("Api.isSuspectedApiCall", () => {
  it("accepts an unchecked action-row checkbox, so a two-way entry can deselect", () => {
    expect(Api.isSuspectedApiCall(actionRowEdit(twoWayColIndex, "FALSE"))).toBe(
      true,
    );
    expect(Api.isSuspectedApiCall(actionRowEdit(twoWayColIndex, "TRUE"))).toBe(
      true,
    );
  });
  it("ignores an action-row edit that isn't a checkbox", () => {
    expect(
      Api.isSuspectedApiCall(actionRowEdit(twoWayColIndex, "some text")),
    ).toBe(false);
  });
  it("ignores a checkbox on the header row or a data row", () => {
    const { tableHeaderRowIndex, topDataRowIndex } = expectedSheetLayout;
    [tableHeaderRowIndex, topDataRowIndex].forEach((rowIndexBase0) => {
      expect(
        Api.isSuspectedApiCall({
          ...actionRowEdit(twoWayColIndex, "TRUE"),
          rowIndexBase0,
        }),
      ).toBe(false);
    });
  });
  it("suspects the header zone's action row in any column, since only a fetch knows where the Table sits", () => {
    const columnCount = getSheetColumnNames("runItem").length;
    expect(Api.isSuspectedApiCall(actionRowEdit(columnCount, "TRUE"))).toBe(
      true,
    );
  });
  it("ignores a sheet with no managed Table", () => {
    expect(
      Api.isSuspectedApiCall({
        ...actionRowEdit(twoWayColIndex, "TRUE"),
        sheetGid: -1,
      }),
    ).toBe(false);
  });
});

describe("Api.handleSheetEdit, a tick the trigger doesn't match", () => {
  it("costs no fetch for a data checkbox", () => {
    const { getByDataFilterCalls } = stubRunItemSheet();
    const calls: string[] = [];
    const installSource = vi.fn();
    const edit = {
      ...actionRowEdit(twoWayColIndex, "TRUE"),
      rowIndexBase0: expectedSheetLayout.topDataRowIndex,
    };
    Api.handleSheetEdit(
      { configs, endpoints: trackingEndpoints(calls) },
      edit,
      installSource,
    );
    Api.init(trackingEndpoints(calls)).handleSheetEdit(edit);

    expect(calls).toEqual([]);
    expect(installSource).not.toHaveBeenCalled();
    expect(getByDataFilterCalls).toHaveLength(0);
  });

  it("runs nothing for a tick past the Table's columns, after the one header zone fetch", () => {
    const { getByDataFilterCalls } = stubRunItemSheet();
    const calls: string[] = [];
    const pastColumns = getSheetColumnNames("runItem").length;

    Api.init(trackingEndpoints(calls)).handleSheetEdit(
      actionRowEdit(pastColumns, "TRUE"),
    );

    expect(calls).toEqual([]);
    expect(getByDataFilterCalls).toHaveLength(1);
  });
});

describe("Api.handleSheetEdit, a Table moved within the header zone", () => {
  const movedStartColIndex = 1;

  function stubMovedRunItemSheet() {
    const shifted = (cells: FakeCell[]): FakeCell[] => [null, ...cells];
    return stubSheetsService({
      sheets: [
        {
          sheetId: runItemGid,
          title: "Run item",
          rows: buildGridRows({
            0: shifted(columnIds),
            3: shifted(["ID", "Selected", "Result", ""]),
            4: shifted(["r:rit:row4", false, "", ""]),
          }),
          table: {
            startColumnIndex: movedStartColIndex,
            endColumnIndex: movedStartColIndex + columnIds.length,
            endRowIndex: 5,
          },
        },
      ],
    });
  }

  it("runs the entry ticked on its live action row, without regenerating", () => {
    stubMovedRunItemSheet();
    const calls: string[] = [];

    Api.init(trackingEndpoints(calls)).handleSheetEdit(
      actionRowEdit(movedStartColIndex + buttonColIndex, "TRUE"),
    );

    expect(calls).toEqual(["button"]);
  });

  it("dispatches in two round trips, one read carrying the header zone and one write", () => {
    const { getByDataFilterCalls, getCalls, batchUpdateCount } =
      stubMovedRunItemSheet();

    Api.init(trackingEndpoints([])).handleSheetEdit(
      actionRowEdit(movedStartColIndex + buttonColIndex, "TRUE"),
    );

    expect(getByDataFilterCalls).toHaveLength(1);
    expect(getCalls).toEqual([]);
    expect(batchUpdateCount()).toBe(1);
  });

  it("ignores a tick on the action row left of where the Table now starts", () => {
    stubMovedRunItemSheet();
    const calls: string[] = [];

    Api.init(trackingEndpoints(calls)).handleSheetEdit(
      actionRowEdit(0, "TRUE"),
    );

    expect(calls).toEqual([]);
  });

  it("ignores a tick below the header zone, without a fetch", () => {
    const { getByDataFilterCalls } = stubMovedRunItemSheet();
    const calls: string[] = [];

    Api.init(trackingEndpoints(calls)).handleSheetEdit({
      ...actionRowEdit(movedStartColIndex + buttonColIndex, "TRUE"),
      rowIndexBase0: SheetIndex.row(5),
    });

    expect(calls).toEqual([]);
    expect(getByDataFilterCalls).toHaveLength(0);
  });
});

describe("Api.handleSheetEdit, a Table that shares its sheet", () => {
  it("runs the entry ticked on its own action row, beside another Table", () => {
    const calls: string[] = [];
    stubSheetsService({ sheets: [layoutSheet()] });
    const endpoints: Endpoints = {
      layoutRight_amount: {
        action: () => {
          calls.push("right amount");
        },
      },
    };

    Api.init(endpoints).handleSheetEdit({
      sheetGid: layoutGid,
      rowIndexBase0: SheetIndex.row(
        layoutOrigins.layoutRight.startRowIndex - 1,
      ),
      colIndexBase0: SheetIndex.col(
        layoutOrigins.layoutRight.startColumnIndex + 1,
      ),
      value: "TRUE",
    });

    expect(calls).toEqual(["right amount"]);
  });

  it("runs the entry ticked on one Table while its neighbour is missing", () => {
    const calls: string[] = [];
    const sheet = layoutSheet();
    stubSheetsService({
      sheets: [
        {
          ...sheet,
          tables: sheet.tables?.filter(
            (table) => table.tableId !== layoutTableId("layoutRight"),
          ),
        },
      ],
    });
    const endpoints: Endpoints = {
      layoutLeft_amount: {
        action: () => {
          calls.push("left amount");
        },
      },
    };

    Api.init(endpoints).handleSheetEdit({
      sheetGid: layoutGid,
      rowIndexBase0: SheetIndex.row(layoutOrigins.layoutLeft.startRowIndex - 1),
      colIndexBase0: SheetIndex.col(
        layoutOrigins.layoutLeft.startColumnIndex + 1,
      ),
      value: "TRUE",
    });

    expect(calls).toEqual(["left amount"]);
  });

  it("stops on the ticked Table when its column ID row holds another Table's IDs", () => {
    stubSheetsService({
      sheets: [
        layoutSheet({
          layoutRight: { headRows: { 3: ["c:lyl:entry", "c:lyl:amount"] } },
        }),
      ],
    });

    expect(() =>
      Api.init({}).handleSheetEdit({
        sheetGid: layoutGid,
        rowIndexBase0: SheetIndex.row(
          layoutOrigins.layoutRight.startRowIndex - 1,
        ),
        colIndexBase0: SheetIndex.col(
          layoutOrigins.layoutRight.startColumnIndex + 1,
        ),
        value: "TRUE",
      }),
    ).toThrowError(/needs its own "lyr" column IDs/);
  });
});

describe("Api.handleSheetEdit, endpoint dispatch", () => {
  it("runs the entry registered under the edited column's full name", () => {
    const calls: string[] = [];
    stubRunItemSheet();

    Api.init(trackingEndpoints(calls)).handleSheetEdit(
      actionRowEdit(buttonColIndex, "TRUE"),
    );

    expect(calls).toEqual(["button"]);
  });

  it("does nothing for a column with no registered entry", () => {
    const calls: string[] = [];
    const { grid } = stubRunItemSheet(tickedAt(idColIndex));
    const before = grid.sheet(runItemGid).rows();

    Api.init(trackingEndpoints(calls)).handleSheetEdit(
      actionRowEdit(idColIndex, "TRUE"),
    );

    expect(calls).toEqual([]);
    expect(grid.sheet(runItemGid).rows()).toEqual(before);
  });

  it("does nothing for a table column that has no column id yet", () => {
    const calls: string[] = [];
    const { grid } = stubRunItemSheet(tickedAt(blankIdColIndex));
    const before = grid.sheet(runItemGid).rows();

    Api.init(trackingEndpoints(calls)).handleSheetEdit(
      actionRowEdit(blankIdColIndex, "TRUE"),
    );

    expect(calls).toEqual([]);
    expect(grid.sheet(runItemGid).rows()).toEqual(before);
  });

  it("ignores an untick for an entry that does not run on uncheck", () => {
    const calls: string[] = [];
    stubRunItemSheet();

    Api.init(trackingEndpoints(calls)).handleSheetEdit(
      actionRowEdit(buttonColIndex, "FALSE"),
    );

    expect(calls).toEqual([]);
  });

  it("runs an entry that declares runOnUncheck on both tick and untick", () => {
    const calls: string[] = [];
    stubRunItemSheet();
    const api = Api.init(trackingEndpoints(calls));

    api.handleSheetEdit(actionRowEdit(twoWayColIndex, "TRUE"));
    api.handleSheetEdit(actionRowEdit(twoWayColIndex, "FALSE"));

    expect(calls).toEqual(["twoWay:true", "twoWay:false"]);
  });
});

describe("Api.handleSheetEdit, the entry checkbox", () => {
  it("clears a button's checkbox so it is ready for the next click", () => {
    const { grid } = stubRunItemSheet(tickedAt(buttonColIndex));

    Api.init(trackingEndpoints([])).handleSheetEdit(
      actionRowEdit(buttonColIndex, "TRUE"),
    );

    expect(actionRowCells(grid)).toEqual([null, null, false, null]);
  });

  it("leaves a two-way entry's checkbox where the operator put it", () => {
    const { grid } = stubRunItemSheet(tickedAt(twoWayColIndex));

    Api.init(trackingEndpoints([])).handleSheetEdit(
      actionRowEdit(twoWayColIndex, "TRUE"),
    );

    expect(actionRowCells(grid)).toEqual([null, true, null, null]);
  });

  it("costs one read and one write for an entry that reports nothing", () => {
    const { batchUpdateCount, getByDataFilterCalls } = stubRunItemSheet();

    Api.init(trackingEndpoints([])).handleSheetEdit(
      actionRowEdit(buttonColIndex, "TRUE"),
    );

    expect(getByDataFilterCalls).toHaveLength(1);
    expect(batchUpdateCount()).toBe(1);
  });
});

describe("Api.handleSheetEdit, the endpoints it installs", () => {
  it("lets a spreadsheet built afterwards reuse a row holding only a run status", async () => {
    vi.resetModules();
    const fresh = await import("./Api");
    const fake = await import("../testSupport/fakeSheetsService");
    const { SpreadsheetNamed } =
      await import("../04_SpreadsheetNamed/SpreadsheetNamed");
    const runStatusColumnId = getColumnTraitByName(
      "runItem",
      "runStatus",
      "columnId",
    );
    const { grid } = fake.stubSheetsService({
      sheets: [
        {
          sheetId: runItemGid,
          title: "Run item",
          rows: buildGridRows({
            0: [
              getColumnTraitByName("runItem", "id", "columnId"),
              getColumnTraitByName("runItem", "result", "columnId"),
              runStatusColumnId,
            ],
            4: [null, null, "Succeeded"],
          }),
          table: {
            endRowIndex: 5,
            tableId: getTableTraitByName("runItem", "tableId"),
          },
        },
      ],
    });
    fresh.Api.handleSheetEdit(
      {
        configs,
        endpoints: {
          runItem_selected: { action: vi.fn(), runStatus: "runStatus" },
        },
      },
      actionRowEdit(idColIndex, "TRUE"),
      vi.fn(),
    );

    const ss = SpreadsheetNamed.init();
    const sheet = ss.table("runItem");
    sheet.row(4).prepFetchFull();
    ss.fetchAllPrepped();
    sheet.appendRowWithVals({ result: "appended" });
    ss.batchUpdateGSheets();

    expect(grid.sheet(runItemGid).values({ startRowIndex: 4 })).toEqual([
      [expect.stringMatching(/^r:rit:/), "appended", ""],
    ]);
  });
});
