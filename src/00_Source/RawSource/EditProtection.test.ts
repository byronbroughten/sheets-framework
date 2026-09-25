import { describe, expect, it } from "vitest";

import {
  assertNotType,
  assertType,
  type IsExactly,
} from "../../testSupport/typeAssertions";
import {
  type EditLockDeclaration,
  type EditWarningDeclaration,
  protectionRangeEqual,
  type WholeSheetEditLockDeclaration,
  type WholeSheetEditWarningDeclaration,
} from "./EditProtection";

type HasUnprotectedRanges<T> = "unprotectedRanges" extends keyof T
  ? true
  : false;

describe("EditProtection identity", () => {
  it("makes unprotected ranges unrepresentable on a non-whole-sheet declaration", () => {
    assertType<IsExactly<HasUnprotectedRanges<EditWarningDeclaration>, false>>(
      true,
    );
    assertType<IsExactly<HasUnprotectedRanges<EditLockDeclaration>, false>>(
      true,
    );
    assertType<
      IsExactly<HasUnprotectedRanges<WholeSheetEditWarningDeclaration>, true>
    >(true);
    assertType<
      IsExactly<HasUnprotectedRanges<WholeSheetEditLockDeclaration>, true>
    >(true);
    assertNotType<
      IsExactly<EditWarningDeclaration, WholeSheetEditWarningDeclaration>
    >(false);
  });

  it("does not treat a whole-column range as a whole-sheet range", () => {
    expect(
      protectionRangeEqual(
        {
          sheetId: 1,
          startRowIndex: 0,
          startColumnIndex: 4,
          endColumnIndex: 5,
        },
        { sheetId: 1 },
      ),
    ).toBe(false);
  });

  it("does not treat whole-column ranges on different columns as equal", () => {
    expect(
      protectionRangeEqual(
        {
          sheetId: 1,
          startRowIndex: 0,
          startColumnIndex: 4,
          endColumnIndex: 5,
        },
        {
          sheetId: 1,
          startRowIndex: 0,
          startColumnIndex: 5,
          endColumnIndex: 6,
        },
      ),
    ).toBe(false);
  });
});
