import { describe, expect, it } from "vitest";

import { assertType, type IsExactly } from "../testSupport/typeAssertions";
import type { SerialDate } from "./SerialDate";
import { Val } from "./Val";

// The calendar rules are SerialDate's; this only proves the facade delegates to it.
describe("Val.validate.date", () => {
  it("takes a whole-day serial and rejects one carrying a time of day", () => {
    expect(Val.validate.date(45292)).toBe(45292);
    expect(() => Val.validate.date(45292.5)).toThrowError(/is not a date/);
    expect(() => Val.validate.date(new Date())).toThrowError(/is not a date/);
  });

  it("hands back a SerialDate rather than a plain number", () => {
    assertType<IsExactly<ReturnType<typeof Val.validate.date>, SerialDate>>(
      true,
    );
  });

  it("narrows an unknown through Val.is.date", () => {
    expect(Val.is.date(45292)).toBe(true);
    expect(Val.is.date(new Date())).toBe(false);
  });
});
