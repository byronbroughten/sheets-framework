import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { getTableTraitByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { stubLogger } from "../testSupport/fakeAppsScriptGlobals";
import {
  buildGridRows,
  type FakeSheetProperties,
} from "../testSupport/fakeSheetsService";
import { fakeTableSheet } from "../testSupport/fakeSheetsService/fakeTableSheet";
import { SyncCoordinator } from "./SyncCoordinator";
import {
  columnsByHeader,
  type GridRange,
  seedFixture,
  sendWithNoAlterations,
} from "./syncCoordinatorTestSupport";

beforeEach(() => {
  stubLogger();
});

afterEach(() => {
  vi.restoreAllMocks();
});

describe("SyncCoordinator conditional-format rule read", () => {
  const rulesGid = 777000222;
  const widgetGid = 777000333;
  // A Table the previous configs grant, so its rules go under its Table key.
  const widgetTableId = getTableTraitByName("log", "tableId");
  const separateRuleGet = "sheets(properties(sheetId),conditionalFormats)";

  function backgroundRule(
    range: GridRange,
  ): GoogleAppsScript.Sheets.Schema.ConditionalFormatRule {
    return {
      ranges: [range],
      booleanRule: {
        condition: {
          type: "CUSTOM_FORMULA",
          values: [{ userEnteredValue: "=$B5>1" }],
        },
        format: { backgroundColor: { red: 1, green: 0, blue: 0 } },
      },
    };
  }

  function rulesSheet(): FakeSheetProperties {
    return {
      sheetId: rulesGid,
      title: "Rules",
      rows: buildGridRows({}),
      conditionalFormats: [
        backgroundRule({
          sheetId: rulesGid,
          startRowIndex: 4,
          endRowIndex: 8,
          startColumnIndex: 1,
          endColumnIndex: 2,
        }),
      ],
    };
  }

  it("brings every sheet's rules into Raw on the regen's plain get, with no separate rule get", () => {
    const service = seedFixture({ extraSheets: [rulesSheet()] });
    const coordinator = SyncCoordinator.init();

    sendWithNoAlterations(coordinator);

    expect(coordinator.ss.raw.sheet(rulesGid).conditionalFormatRules()).toEqual(
      [
        {
          kind: "boolean",
          ranges: [
            {
              sheetId: rulesGid,
              startRowIndex: 4,
              endRowIndex: 8,
              startColumnIndex: 1,
              endColumnIndex: 2,
            },
          ],
          condition: { type: "CUSTOM_FORMULA", formula: "=$B5>1" },
          format: { backgroundColor: { red: 1, green: 0, blue: 0 } },
        },
      ],
    );
    expect(service.getCalls.map(({ fields }) => fields)).not.toContain(
      separateRuleGet,
    );
  });

  it("writes a Table's rules under its Table key and a tab's under its tab name, each by its rule handle", () => {
    seedFixture({
      extraSheets: [
        rulesSheet(),
        {
          ...fakeTableSheet.build({
            sheetId: widgetGid,
            title: "Widget",
            tableId: widgetTableId,
            name: "Widget",
            columnConfigs: columnsByHeader(["Name", "Score"]),
            columnNames: ["Name", "Score"],
            bodyRows: [{}],
          }),
          conditionalFormats: [
            backgroundRule({
              sheetId: widgetGid,
              startRowIndex: 4,
              endRowIndex: 5,
              startColumnIndex: 1,
              endColumnIndex: 2,
            }),
          ],
        },
      ],
    });

    const { conditionalFormatConfigs } = sendWithNoAlterations();

    expect(conditionalFormatConfigs).toContain(
      "export const conditionalFormatConfigs = makeConditionalFormatConfigs({",
    );
    expect(conditionalFormatConfigs).toContain(
      `  "tables": {\n    "widget": {\n      "score | =\${@score}>1 | background #ff0000": { "kind": "boolean", "tableId": "${widgetTableId}", "columns": ["score"], "condition": {"type":"CUSTOM_FORMULA","formula":"=\${@score}>1"}, "format": {"backgroundColor":"#ff0000"} }\n    }\n  },`,
    );
    expect(conditionalFormatConfigs).toContain(
      `  "tabs": {\n    "rules": {\n      "B5:B8 | =$B5>1 | background #ff0000": { "kind": "boolean", "sheetGid": ${rulesGid}, "ranges": ["B5:B8"], "condition": {"type":"CUSTOM_FORMULA","formula":"=$B5>1"}, "format": {"backgroundColor":"#ff0000"} }\n    }\n  }\n});`,
    );
  });
});
