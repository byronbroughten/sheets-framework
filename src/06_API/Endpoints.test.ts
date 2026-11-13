import { describe, it } from "vitest";

import type {
  ColumnFullName,
  ColumnNameFiltered,
} from "../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import type { TableNameSimple } from "../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { assertType, type IsExactly } from "../testSupport/typeAssertions";
import type { Endpoint, Endpoints, EndpointsAll } from "./Endpoints";

type SelectorColumnOf<TN extends TableNameSimple> = NonNullable<
  Endpoint<TN>["selector"]
>["column"];
type TimeLastRanOf<TN extends TableNameSimple> = NonNullable<
  Endpoint<TN>["timeLastRan"]
>;

describe("Endpoint's column parameters", () => {
  it("resolve to the entry column's own sheet", () => {
    assertType<IsExactly<SelectorColumnOf<"tableConfig">, "letApiAccess">>(
      true,
    );
    assertType<
      IsExactly<
        TimeLastRanOf<"spreadsheetConfig">,
        ColumnNameFiltered<"spreadsheetConfig", "string", false>
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
        SelectorColumnOf<TableNameSimple>,
        ColumnNameFiltered<TableNameSimple, "checkbox", false>
      >
    >(true);
    assertType<
      IsExactly<
        Extract<SelectorColumnOf<TableNameSimple>, "selected">,
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
  it("rejects an ID column, whose entry the framework derives", () => {
    assertType<IsExactly<Extract<keyof Endpoints, "runItem_id">, never>>(true);
    assertType<IsExactly<Extract<keyof Endpoints, `${string}_id`>, never>>(
      true,
    );
  });
  it("keeps every other column, a config sheet's included", () => {
    assertType<
      IsExactly<keyof Endpoints, Exclude<ColumnFullName, `${string}_id`>>
    >(true);
    assertType<
      IsExactly<
        Extract<keyof Endpoints, "tableConfig_letApiAccess">,
        "tableConfig_letApiAccess"
      >
    >(true);
    assertType<
      IsExactly<Extract<keyof Endpoints, "runItem_result">, "runItem_result">
    >(true);
  });
  it("types each entry by its own column's sheet", () => {
    assertType<
      IsExactly<Endpoints["runItem_result"], Endpoint<"runItem"> | undefined>
    >(true);
  });
});
