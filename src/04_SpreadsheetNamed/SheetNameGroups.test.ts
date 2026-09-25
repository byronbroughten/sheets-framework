import { describe, expect, it } from "vitest";

import { assertType, type IsExactly } from "../testSupport/typeAssertions";
import {
  isInTnGroup,
  type SheetNameWithIdAndNameColumn,
  type SheetNameWithIdColumn,
  type SheetNameWithNameColumn,
} from "./SheetNameGroups";

describe("SheetNameWithIdColumn", () => {
  it("includes a sheet that declares an id column and excludes one that doesn't", () => {
    assertType<IsExactly<Extract<SheetNameWithIdColumn, "item">, "item">>(true);
    assertType<
      IsExactly<Extract<SheetNameWithIdColumn, "spreadsheetConfig">, never>
    >(true);
  });

  it("agrees at runtime with the type it is derived from", () => {
    expect(isInTnGroup("hasIdColumn", "item")).toBe(true);
    expect(isInTnGroup("hasIdColumn", "spreadsheetConfig")).toBe(false);
  });
});

describe("SheetNameWithNameColumn", () => {
  it("includes a sheet whose header row has the Name header and excludes one that doesn't", () => {
    assertType<IsExactly<Extract<SheetNameWithNameColumn, "item">, "item">>(
      true,
    );
    assertType<
      IsExactly<Extract<SheetNameWithNameColumn, "spreadsheetConfig">, never>
    >(true);
  });
});

describe("SheetNameWithIdAndNameColumn", () => {
  it("excludes a sheet with an id column but no name column, and one with neither", () => {
    assertType<
      IsExactly<
        Extract<SheetNameWithIdAndNameColumn, "item" | "valueTypes" | "log">,
        "item"
      >
    >(true);
  });
});
