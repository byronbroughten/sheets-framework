import { describe, expect, it } from "vitest";

import { assertType, type IsExactly } from "../testSupport/typeAssertions";
import {
  type RemoveFirstN,
  type SentenceToCamelCase,
  Str,
  type TakeFirstN,
} from "./Str";

describe("Str.combineStrings", () => {
  it("joins the two strings", () => {
    expect(Str.combineStrings("row", "Id")).toBe("rowId");
  });
});

describe("Str.removeFirstN", () => {
  it("drops the first n characters", () => {
    expect(Str.removeFirstN("rowId", 3)).toBe("Id");
    expect(Str.removeFirstN("ab", 5)).toBe("");
  });
});

describe("Str.takeFirstN", () => {
  it("keeps only the first n characters", () => {
    expect(Str.takeFirstN("rowId", 3)).toBe("row");
    expect(Str.takeFirstN("ab", 5)).toBe("ab");
  });
});

describe("Str.sentenceToCamelCase", () => {
  it("camelCases a header sentence", () => {
    expect(Str.sentenceToCamelCase("Column ID")).toBe("columnId");
  });

  it("splits on runs of punctuation and trims the ends", () => {
    expect(Str.sentenceToCamelCase("  Rent / month (USD) ")).toBe(
      "rentMonthUsd",
    );
  });

  it("removes apostrophes rather than splitting on them", () => {
    expect(Str.sentenceToCamelCase("Tenant's Name")).toBe("tenantsName");
    expect(Str.sentenceToCamelCase("Tenant’s Name")).toBe("tenantsName");
  });
});

describe("RemoveFirstN", () => {
  it("drops the first n characters, leaving an empty string past the end", () => {
    assertType<IsExactly<RemoveFirstN<"rowId", 3>, "Id">>(true);
    assertType<IsExactly<RemoveFirstN<"ab", 5>, "">>(true);
  });
});

describe("TakeFirstN", () => {
  it("keeps the first n characters, or the whole string when it is shorter", () => {
    assertType<IsExactly<TakeFirstN<"rowId", 3>, "row">>(true);
    assertType<IsExactly<TakeFirstN<"ab", 5>, "ab">>(true);
  });
});

describe("SentenceToCamelCase", () => {
  it("mirrors Str.sentenceToCamelCase at the type level", () => {
    assertType<IsExactly<SentenceToCamelCase<"Column ID">, "columnId">>(true);
    assertType<
      IsExactly<SentenceToCamelCase<"  Rent / month (USD) ">, "rentMonthUsd">
    >(true);
    assertType<IsExactly<SentenceToCamelCase<"Tenant's Name">, "tenantsName">>(
      true,
    );
  });
});
