import { describe, expect, it } from "vitest";

import { assertType, type IsExactly } from "../testSupport/typeAssertions";
import { Obj } from "./Obj";

describe("Obj.toKeyedMap", () => {
  it("keys each entry by its id field and stamps its outer key as the name", () => {
    const map = Obj.toKeyedMap(
      { first: { id: 7, label: "a" }, second: { id: 9, label: "b" } },
      "id",
    );
    expect([...map.entries()]).toEqual([
      [7, { id: 7, label: "a", name: "first" }],
      [9, { id: 9, label: "b", name: "second" }],
    ]);
  });

  it("stamps the outer key under a chosen name field", () => {
    const map = Obj.toKeyedMap({ first: { id: 7 } }, "id", "key");
    expect(map.get(7)).toEqual({ id: 7, key: "first" });
  });
});

describe("Obj.pushByKey", () => {
  it("starts an array for a missing key and appends to an existing one", () => {
    const obj: Record<string, number[]> = { a: [1] };
    Obj.pushByKey(obj, "a", 2);
    Obj.pushByKey(obj, "b", 3);
    expect(obj).toEqual({ a: [1, 2], b: [3] });
  });
});

describe("Obj.keyByValue", () => {
  it("returns the first key holding the value", () => {
    expect(Obj.keyByValue({ a: 1, b: 2, c: 2 }, 2)).toBe("b");
  });

  it("throws when no key holds the value", () => {
    const obj: Record<string, number> = { a: 1 };
    expect(() => Obj.keyByValue(obj, 5)).toThrowError(
      "Value not found in object",
    );
  });
});

describe("Obj.isEmpty", () => {
  it("says whether the object has no own keys", () => {
    expect(Obj.isEmpty({})).toBe(true);
    expect(Obj.isEmpty({ a: undefined })).toBe(false);
  });
});

describe("Obj.isKey", () => {
  it("says whether the value is one of the object's keys", () => {
    expect(Obj.isKey({ a: 1 }, "a")).toBe(true);
    expect(Obj.isKey({ a: 1 }, "b")).toBe(false);
  });
});

describe("Obj.stringifyEqual", () => {
  it("compares by JSON, so key order matters", () => {
    expect(Obj.stringifyEqual({ a: [1] }, { a: [1] })).toBe(true);
    expect(Obj.stringifyEqual({ a: 1, b: 2 }, { b: 2, a: 1 })).toBe(false);
  });
});

describe("Obj.isObjToAny", () => {
  it("accepts any non-null object, arrays included", () => {
    expect(Obj.isObjToAny({})).toBe(true);
    expect(Obj.isObjToAny([])).toBe(true);
    expect(Obj.isObjToAny(null)).toBe(false);
    expect(Obj.isObjToAny("a")).toBe(false);
  });
});

describe("Obj.isObjToRecord", () => {
  it("accepts any non-null object, arrays included", () => {
    expect(Obj.isObjToRecord({ a: 1 })).toBe(true);
    expect(Obj.isObjToRecord([])).toBe(true);
    expect(Obj.isObjToRecord(undefined)).toBe(false);
    expect(Obj.isObjToRecord(3)).toBe(false);
  });
});

describe("Obj.pick", () => {
  it("copies the listed keys that are present and skips the rest", () => {
    const obj: { a: number; b: number; c?: number } = { a: 1, b: 2 };
    expect(Obj.pick(obj, ["a", "c"])).toEqual({ a: 1 });
  });
});

describe("Obj.validatePick", () => {
  it("copies the listed keys after validating each value", () => {
    expect(
      Obj.validatePick({ a: "x", b: "y", c: 3 }, "string", "a", "b"),
    ).toEqual({ a: "x", b: "y" });
  });

  it("throws on a value of the wrong kind", () => {
    expect(() => Obj.validatePick({ a: 1 }, "string", "a")).toThrowError(
      /is not a string/,
    );
  });
});

describe("Obj.strictPick", () => {
  it("copies the listed keys", () => {
    expect(Obj.strictPick({ a: 1, b: 2 }, ["b"])).toEqual({ b: 2 });
  });

  it("throws on a listed key that is absent", () => {
    const obj: { a: number; b?: number } = { a: 1 };
    expect(() => Obj.strictPick(obj, ["b"])).toThrowError(
      "Key b not in object",
    );
  });
});

describe("Obj.pickStartsWith", () => {
  it("keeps only the keys that start with the prefix", () => {
    expect(Obj.pickStartsWith({ rowA: 1, rowB: 2, colA: 3 }, "row")).toEqual({
      rowA: 1,
      rowB: 2,
    });
  });
});

describe("Obj.removeFirstNFromKeys", () => {
  it("drops the first n characters of every key", () => {
    expect(Obj.removeFirstNFromKeys({ rowA: 1, rowB: 2 }, 3)).toEqual({
      A: 1,
      B: 2,
    });
  });
});

describe("Obj.strictOmit", () => {
  it("copies every key but the omitted ones", () => {
    expect(Obj.strictOmit({ a: 1, b: 2, c: 3 }, "a", "c")).toEqual({ b: 2 });
  });
});

describe("Obj.keys", () => {
  it("lists the object's own keys", () => {
    expect(Obj.keys({ a: 1, b: 2 })).toEqual(["a", "b"]);
  });

  it("is typed as an array of keys, not a one-item tuple", () => {
    assertType<
      IsExactly<ReturnType<typeof Obj.keys<{ a: 1; b: 2 }>>, ("a" | "b")[]>
    >(true);
  });
});

describe("Obj.stringKeys", () => {
  it("lists the object's own keys", () => {
    expect(Obj.stringKeys({ a: 1, b: 2 })).toEqual(["a", "b"]);
  });
});

describe("Obj.values", () => {
  it("lists the object's own values", () => {
    expect(Obj.values({ a: 1, b: "x" })).toEqual([1, "x"]);
  });
});

describe("Obj.entries", () => {
  it("lists the object's own key-value pairs", () => {
    expect(Obj.entries({ a: 1, b: "x" })).toEqual([
      ["a", 1],
      ["b", "x"],
    ]);
  });
});

describe("Obj.mapValues", () => {
  it("maps each value with its key and keeps the keys", () => {
    expect(
      Obj.mapValues({ a: 1, b: 2 }, (value, key) => `${String(key)}${value}`),
    ).toEqual({ a: "a1", b: "b2" });
  });
});

describe("Obj.propKeysOfValue", () => {
  it("lists every key holding the value", () => {
    expect(Obj.propKeysOfValue({ a: 1, b: 2, c: 1 }, 1)).toEqual(["a", "c"]);
  });
});

describe("Obj.invert", () => {
  it("swaps keys and values, the last key winning a shared value", () => {
    expect(Obj.invert({ a: "x", b: "y", c: "x" })).toEqual({ x: "c", y: "b" });
  });
});

describe("Obj.merge", () => {
  it("copies both objects into a new one, the second winning a shared key", () => {
    const a = { a: 1, b: 1 };
    expect(Obj.merge(a, { b: 2, c: 3 })).toEqual({ a: 1, b: 2, c: 3 });
    expect(a).toEqual({ a: 1, b: 1 });
  });
});

describe("Obj.spread", () => {
  it("copies every object into a new one, later ones winning a shared key", () => {
    expect(Obj.spread({ a: 1 }, { a: 2, b: 2 }, { c: 3 })).toEqual({
      a: 2,
      b: 2,
      c: 3,
    });
  });
});
