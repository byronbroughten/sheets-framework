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

  it("gives one key for sentence, title, snake and camel spellings", () => {
    ["add occ charge", "Add Occ Charge", "add_occ_charge", "addOccCharge"].forEach(
      (spelling) => {
        expect(Str.sentenceToCamelCase(spelling)).toBe("addOccCharge");
      },
    );
  });

  it("ends an acronym before the capital that starts the next word", () => {
    expect(Str.sentenceToCamelCase("parseHTMLTable")).toBe("parseHtmlTable");
    expect(Str.sentenceToCamelCase("Unit ID")).toBe("unitId");
  });

  it("keeps a mixed-case word whole inside a sentence", () => {
    expect(Str.sentenceToCamelCase("CapEx budget")).toBe("capexBudget");
  });

  it("gives back a camelCase key unchanged", () => {
    ["addOccCharge", "unit2Rent", "columnId", "capexBudget"].forEach((key) => {
      expect(Str.sentenceToCamelCase(key)).toBe(key);
    });
  });

  it("keeps a plural acronym as one word", () => {
    expect(Str.sentenceToCamelCase("Unit IDs")).toBe("unitIds");
    expect(Str.sentenceToCamelCase("unitIDsCount")).toBe("unitIdsCount");
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

  it("gives one key for sentence, title, snake and camel spellings", () => {
    assertType<IsExactly<SentenceToCamelCase<"add occ charge">, "addOccCharge">>(
      true,
    );
    assertType<IsExactly<SentenceToCamelCase<"Add Occ Charge">, "addOccCharge">>(
      true,
    );
    assertType<IsExactly<SentenceToCamelCase<"add_occ_charge">, "addOccCharge">>(
      true,
    );
    assertType<IsExactly<SentenceToCamelCase<"addOccCharge">, "addOccCharge">>(
      true,
    );
  });

  it("splits a token on case like the runtime, and a sentence only on punctuation", () => {
    assertType<
      IsExactly<SentenceToCamelCase<"parseHTMLTable">, "parseHtmlTable">
    >(true);
    assertType<IsExactly<SentenceToCamelCase<"unitIDsCount">, "unitIdsCount">>(
      true,
    );
    assertType<IsExactly<SentenceToCamelCase<" unit2Rent ">, "unit2Rent">>(
      true,
    );
    assertType<IsExactly<SentenceToCamelCase<"CapEx budget">, "capexBudget">>(
      true,
    );
    assertType<IsExactly<SentenceToCamelCase<"Unit IDs">, "unitIds">>(true);
  });

  it("widens to string for a string that isn't a literal", () => {
    assertType<IsExactly<SentenceToCamelCase<string>, string>>(true);
  });
});
