import { describe, it } from "vitest";

import type {
  ColumnFullName,
  ColumnNameFiltered,
} from "../01_SpreadsheetSchema/columnConfigsTypes";
import type { FloorTabName } from "../01_SpreadsheetSchema/configSheetFloorSeed";
import type { SheetNameSimple } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import { assertType, type IsExactly } from "../testSupport/typeAssertions";
import type { Endpoint, Endpoints, EndpointsAll } from "./Endpoints";

type SelectorColumnOf<SN extends SheetNameSimple> = NonNullable<
  Endpoint<SN>["selector"]
>["column"];
type TimeLastRanOf<SN extends SheetNameSimple> = NonNullable<
  Endpoint<SN>["timeLastRan"]
>;

describe("Endpoint's column parameters", () => {
  it("resolve to the entry column's own sheet", () => {
    assertType<IsExactly<SelectorColumnOf<"sheetConfig">, "letApiAccess">>(
      true,
    );
    assertType<
      IsExactly<
        TimeLastRanOf<"spreadsheetConfig">,
        | "idDelimiter"
        | "idHeader"
        | "nameHeader"
        | "tableMenuSpace"
        | "fillRowIdsTimeLastRan"
        | "fillRowIdsRunStatus"
        | "syncConfigSheetRowsTimeLastRan"
        | "syncConfigSheetRowsRunStatus"
      >
    >(true);
  });

  it("are filtered to the value type each one needs", () => {
    assertType<
      IsExactly<
        SelectorColumnOf<"runItem">,
        ColumnNameFiltered<"runItem", "checkbox", false>
      >
    >(true);
    assertType<
      IsExactly<
        TimeLastRanOf<"runItem">,
        ColumnNameFiltered<"runItem", "string", false>
      >
    >(true);
  });
});

describe("Endpoint at the widened sheet name the dispatch boundary uses", () => {
  it("resolves to the cross-sheet union rather than to never", () => {
    assertType<
      IsExactly<
        SelectorColumnOf<SheetNameSimple>,
        ColumnNameFiltered<SheetNameSimple, "checkbox", false>
      >
    >(true);
    assertType<
      IsExactly<
        Extract<SelectorColumnOf<SheetNameSimple>, "selected">,
        "selected"
      >
    >(true);
  });
});

describe("EndpointsAll", () => {
  it("is keyed by every column full name and nothing else", () => {
    assertType<IsExactly<keyof EndpointsAll, ColumnFullName>>(true);
  });
});

describe("Endpoints, the endpoint map", () => {
  it("rejects a column on a config sheet", () => {
    assertType<
      IsExactly<
        Extract<keyof Endpoints, "spreadsheetConfig_fillRowIdsTimeLastRan">,
        never
      >
    >(true);
    assertType<
      IsExactly<Extract<keyof Endpoints, `${FloorTabName}_${string}`>, never>
    >(true);
  });
  it("keeps every column on an app sheet", () => {
    assertType<
      IsExactly<
        keyof Endpoints,
        Exclude<ColumnFullName, `${FloorTabName}_${string}`>
      >
    >(true);
    assertType<
      IsExactly<Extract<keyof Endpoints, "runItem_result">, "runItem_result">
    >(true);
  });
});
