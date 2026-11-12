import { describe, it } from "vitest";

import type { configSheetFloorSeed } from "../01_SpreadsheetSchema/configReaders/configSheetFloorSeed";
import { assertType, type IsExactly } from "../testSupport/typeAssertions";
import type { frameworkEndpoints } from "./frameworkEndpoints";

describe("frameworkEndpoints", () => {
  it("must supply exactly the seeded endpoints but Sync Config, which is gone", () => {
    assertType<
      IsExactly<
        keyof typeof frameworkEndpoints,
        Exclude<
          keyof typeof configSheetFloorSeed.spreadsheetConfig.endpoints,
          "spreadsheetConfig_syncConfigSheetRowsTimeLastRan"
        >
      >
    >(true);
  });
});
