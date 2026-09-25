import { afterEach, describe, expect, it } from "vitest";

import {
  clearSpreadsheetConfigOverlay,
  overlaySpreadsheetConfig,
} from "./spreadsheetConfigTypes";
import { uniformRowLayout } from "./uniformRowLayout";

const validLayout = {
  idDelimiter: ":",
  idHeader: "ID",
  nameHeader: "Name",
  startTableColIndexBase0: 0,
  columnIdRowIdxBase0: 0,
  columnGroupHeadingRowIndexBase0: 1,
  actionRowIndexBase0: 2,
  tableHeaderRowIndexBase0: 3,
} as const;

afterEach(() => {
  clearSpreadsheetConfigOverlay();
});

describe("uniformRowLayout.validate", () => {
  it("throws when two uniform-row indexes share a number", () => {
    expect(() =>
      uniformRowLayout.validate({
        ...validLayout,
        actionRowIndexBase0: 0,
      }),
    ).toThrow(/Column ID row index base 1.*Action row index base 1/);
  });

  it("throws when a uniform-row index lands on the first data row", () => {
    expect(() =>
      uniformRowLayout.validate({
        ...validLayout,
        columnGroupHeadingRowIndexBase0: 4,
      }),
    ).toThrow(/Column group heading row index base 1.*first data row/);
  });
});

describe("overlaySpreadsheetConfig", () => {
  it("throws when overlaying a colliding layout even without a fetch", () => {
    expect(() =>
      overlaySpreadsheetConfig({
        ...validLayout,
        actionRowIndexBase0: 1,
      }),
    ).toThrow(/Column group heading row index base 1.*Action row index base 1/);
  });
});
