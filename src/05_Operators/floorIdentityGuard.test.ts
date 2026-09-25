import { describe, expect, it } from "vitest";

import { installedConfigs } from "../01_SpreadsheetSchema/configRegister";
import type {
  ColumnConfigStored,
  SheetConfigsBase,
} from "../01_SpreadsheetSchema/makeConfigs";
import {
  assertFloorIdentityUnchanged,
  type FloorIdentitySource,
} from "./floorIdentityGuard";

const { columnConfigs } = installedConfigs();

const sheetGidHeader = columnConfigs.sheetConfig.sheetGid.header;

function column(columnId: string): ColumnConfigStored {
  return { ...columnConfigs.sheetConfig.sheetGid, columnId };
}

function source(props: {
  sheetConfigs?: SheetConfigsBase;
  sheetConfigColumnId?: string;
}): FloorIdentitySource {
  return {
    sheetConfigs: props.sheetConfigs ?? {
      sheetConfig: {
        sheetGid: 1,
        idPrefix: "scf",
        hasIdColumn: false,
        hasNameColumn: false,
      },
    },
    columnConfigs:
      props.sheetConfigColumnId === undefined
        ? {}
        : { sheetConfig: { sheetGid: column(props.sheetConfigColumnId) } },
  };
}

describe("assertFloorIdentityUnchanged", () => {
  it("throws naming the sheet, the header, and the previous and new column ID of a floor column", () => {
    expect(() =>
      assertFloorIdentityUnchanged({
        previous: source({ sheetConfigColumnId: "c:scf:aaa" }),
        next: source({ sheetConfigColumnId: "c:scf:bbb" }),
      }),
    ).toThrow(
      `Floor column "${sheetGidHeader}" on "sheetConfig" had column ID "c:scf:aaa" and is now "c:scf:bbb".`,
    );
  });

  it("passes when the floor identities match", () => {
    expect(() =>
      assertFloorIdentityUnchanged({
        previous: source({ sheetConfigColumnId: "c:scf:aaa" }),
        next: source({ sheetConfigColumnId: "c:scf:aaa" }),
      }),
    ).not.toThrow();
  });

  it("skips a floor tab absent from the previous sheet configs", () => {
    expect(() =>
      assertFloorIdentityUnchanged({
        previous: source({ sheetConfigs: {} }),
        next: source({}),
      }),
    ).not.toThrow();
  });

  it("skips a floor column absent from the previous column configs", () => {
    expect(() =>
      assertFloorIdentityUnchanged({
        previous: source({}),
        next: source({ sheetConfigColumnId: "c:scf:bbb" }),
      }),
    ).not.toThrow();
  });
});
