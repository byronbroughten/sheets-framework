import { describe, it } from "vitest";

import type { configSheetFloorSeed } from "../01_SpreadsheetSchema/configSheetFloorSeed";
import { assertType, type IsExactly } from "../testSupport/typeAssertions";
import type { frameworkEndpoints } from "./frameworkEndpoints";

describe("frameworkEndpoints", () => {
  it("must supply exactly the seeded endpoints", () => {
    assertType<
      IsExactly<
        keyof typeof frameworkEndpoints,
        keyof typeof configSheetFloorSeed.spreadsheetConfig.endpoints
      >
    >(true);
  });
});
