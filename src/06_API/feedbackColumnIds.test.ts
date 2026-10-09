import { describe, expect, it, vi } from "vitest";

import { getColumnTraitByName } from "../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import { getTableTraitByName } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { feedbackColumnIdsOf } from "./feedbackColumnIds";

describe("installEndpoints", () => {
  it("throws when the configs aren't installed yet", async () => {
    vi.resetModules();
    const { installEndpoints } = await import("./feedbackColumnIds");
    expect(() => installEndpoints({})).toThrow(
      "Configs have not been installed.",
    );
  });
});

describe("feedbackColumnIdsOf", () => {
  it("keys each Table's feedback column IDs by its tableId", () => {
    const ids = feedbackColumnIdsOf({
      runItem_startTime: {
        action: () => undefined,
        timeLastRan: "startTime",
        runStatus: "runStatus",
      },
    });
    expect(ids).toEqual(
      new Map([
        [
          getTableTraitByName("runItem", "tableId"),
          new Set([
            getColumnTraitByName("runItem", "startTime", "columnId"),
            getColumnTraitByName("runItem", "runStatus", "columnId"),
          ]),
        ],
      ]),
    );
  });
});
