import type { SerialDate } from "@byronbroughten/utils/serial-date";
import { describe, expect, it } from "vitest";

import { assertType, type IsExactly } from "../testSupport/typeAssertions";
import { Val, validationError } from "./Val";

describe("validationError", () => {
  it("builds a plain Error naming the value and what it is not", () => {
    const error = validationError("x", "number");
    expect(error.constructor).toBe(Error);
    expect(error.message).toBe(`value "x" is not a number`);
  });
});

describe("Val.is", () => {
  it("tells each primitive apart", () => {
    expect(Val.is.string("a")).toBe(true);
    expect(Val.is.string(1)).toBe(false);
    expect(Val.is.number(1)).toBe(true);
    expect(Val.is.number("1")).toBe(false);
    expect(Val.is.boolean(false)).toBe(true);
    expect(Val.is.boolean(0)).toBe(false);
  });

  it("accepts only the empty string as emptyString", () => {
    expect(Val.is.emptyString("")).toBe(true);
    expect(Val.is.emptyString(" ")).toBe(false);
    expect(Val.is.emptyString(undefined)).toBe(false);
  });

  it("accepts only a string starting with = as a formula", () => {
    expect(Val.is.formula("=A1+B1")).toBe(true);
    expect(Val.is.formula("A1=B1")).toBe(false);
    expect(Val.is.formula("")).toBe(false);
    expect(Val.is.formula(1)).toBe(false);
  });
});

describe("Val.validate", () => {
  it("hands back a value of the right kind and throws on any other", () => {
    expect(Val.validate.string("a")).toBe("a");
    expect(() => Val.validate.string(1)).toThrowError(
      `value "1" is not a string`,
    );
    expect(Val.validate.number(1)).toBe(1);
    expect(() => Val.validate.number("1")).toThrowError(
      `value "1" is not a number`,
    );
    expect(Val.validate.boolean(false)).toBe(false);
    expect(() => Val.validate.boolean(0)).toThrowError(
      `value "0" is not a boolean`,
    );
  });

  it("takes a number or the empty string for numberOrEmpty", () => {
    expect(Val.validate.numberOrEmpty(1)).toBe(1);
    expect(Val.validate.numberOrEmpty("")).toBe("");
    expect(() => Val.validate.numberOrEmpty("1")).toThrowError(
      `value "1" is not a number or empty string`,
    );
  });
});

describe("Val.assert", () => {
  it("hands back anything but null and undefined, falsy values included", () => {
    expect(Val.assert(0)).toBe(0);
    expect(Val.assert("")).toBe("");
    expect(Val.assert(false)).toBe(false);
  });

  it("throws naming what was not found", () => {
    expect(() => Val.assert(undefined)).toThrowError("Value not found.");
    expect(() => Val.assert(null, "Row")).toThrowError("Row not found.");
  });
});

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
