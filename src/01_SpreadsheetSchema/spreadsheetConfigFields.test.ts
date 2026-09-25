import { describe, it } from "vitest";

import { assertType, type IsExactly } from "../testSupport/typeAssertions";
import type { SentenceToCamelCase } from "../utils/Str";
import type { ColumnName } from "./columnConfigsTypes";
import type {
  spreadsheetConfigIndexHeaders,
  spreadsheetConfigTextHeaders,
} from "./spreadsheetConfigFields";

type PairedHeader =
  | (typeof spreadsheetConfigTextHeaders)[keyof typeof spreadsheetConfigTextHeaders]
  | (typeof spreadsheetConfigIndexHeaders)[keyof typeof spreadsheetConfigIndexHeaders];

describe("spreadsheetConfigFields", () => {
  it("pairs every field with a header that gen:configs names as a Spreadsheet Config column", () => {
    assertType<
      IsExactly<
        Exclude<
          SentenceToCamelCase<PairedHeader>,
          ColumnName<"spreadsheetConfig">
        >,
        never
      >
    >(true);
  });
});
