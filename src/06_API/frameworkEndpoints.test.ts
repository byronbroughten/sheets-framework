import { describe, it } from "vitest";

import type { SheetNameWithIdColumn } from "../04_SpreadsheetNamed/TableNameGroups";
import { assertType, type IsExactly } from "../testSupport/typeAssertions";
import type { Endpoint, EndpointsAll } from "./Endpoints";
import type { FillRowIdsEndpoints } from "./frameworkEndpoints";

describe("FillRowIdsEndpoints", () => {
  it("is keyed by the ID column of every Table with one, and nothing else", () => {
    assertType<
      IsExactly<keyof FillRowIdsEndpoints, `${SheetNameWithIdColumn}_id`>
    >(true);
    assertType<
      IsExactly<Extract<keyof FillRowIdsEndpoints, "runItem_id">, "runItem_id">
    >(true);
  });
  it("leaves out a Table without an ID column", () => {
    assertType<
      IsExactly<
        Extract<keyof FillRowIdsEndpoints, "spreadsheetConfig_id">,
        never
      >
    >(true);
  });
  it("types each entry by its own Table", () => {
    assertType<
      IsExactly<FillRowIdsEndpoints["runItem_id"], Endpoint<"runItem">>
    >(true);
  });
  it("names only columns EndpointsAll keys", () => {
    assertType<
      IsExactly<Exclude<keyof FillRowIdsEndpoints, keyof EndpointsAll>, never>
    >(true);
  });
});
