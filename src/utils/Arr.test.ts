import { describe, expect, it } from "vitest";

import { Arr } from "./Arr";

describe("Arr.indexesFromUntil", () => {
  it("lists every index from the start up to but not including the end", () => {
    expect(Arr.indexesFromUntil(2, 5)).toEqual([2, 3, 4]);
    expect(Arr.indexesFromUntil(3, 3)).toEqual([]);
  });
});

describe("Arr.contiguousRanges", () => {
  it("collapses unsorted indexes into ascending half-open ranges", () => {
    expect(Arr.contiguousRanges([5, 1, 2, 3, 7, 8])).toEqual([
      { startIndex: 1, endIndex: 4 },
      { startIndex: 5, endIndex: 6 },
      { startIndex: 7, endIndex: 9 },
    ]);
    expect(Arr.contiguousRanges([])).toEqual([]);
  });

  it("leaves its input in its original order", () => {
    const indexes = [3, 1, 2];
    Arr.contiguousRanges(indexes);
    expect(indexes).toEqual([3, 1, 2]);
  });
});

describe("Arr.numsInOffsetLength", () => {
  it("lists length numbers counting up from the offset", () => {
    expect(Arr.numsInOffsetLength(4, 3)).toEqual([4, 5, 6]);
    expect(Arr.numsInOffsetLength(4, 0)).toEqual([]);
  });
});

describe("Arr.indexesOf", () => {
  it("lists every index holding the value", () => {
    expect(Arr.indexesOf(["a", "b", "a"], "a")).toEqual([0, 2]);
    expect(Arr.indexesOf(["a"], "z")).toEqual([]);
  });
});

describe("Arr.lastIndex", () => {
  it("is one less than the length, so -1 for an empty array", () => {
    expect(Arr.lastIndex(["a", "b"])).toBe(1);
    expect(Arr.lastIndex([])).toBe(-1);
  });
});

describe("Arr.isLastIndex", () => {
  it("says whether the index is the last one", () => {
    expect(Arr.isLastIndex(["a", "b"], 1)).toBe(true);
    expect(Arr.isLastIndex(["a", "b"], 0)).toBe(false);
  });
});

describe("Arr.indexOrThrow", () => {
  it("returns the index of the first item the finder accepts", () => {
    expect(Arr.indexOrThrow([1, 2, 3, 2], (n) => n === 2)).toBe(1);
  });

  it("throws when no item matches", () => {
    expect(() => Arr.indexOrThrow([1], (n) => n === 2)).toThrowError(
      "Value not found at any index.",
    );
  });
});

describe("Arr.validateIndexOrThrow", () => {
  it("returns true for an index the array has", () => {
    expect(Arr.validateIndexOrThrow(["a", "b"], 1)).toBe(true);
  });

  it("throws for an index past the end", () => {
    expect(() => Arr.validateIndexOrThrow(["a", "b"], 2)).toThrowError(
      "The passed array does not have a value at passed idx 2",
    );
  });
});

describe("Arr.compareForSort", () => {
  it("orders numbers, strings and dates by their natural order", () => {
    expect(Arr.compareForSort(1, 3)).toBeLessThan(0);
    expect(Arr.compareForSort("b", "a")).toBeGreaterThan(0);
    expect(Arr.compareForSort(new Date(5), new Date(5))).toBe(0);
  });

  it("falls back to comparing mixed types as strings", () => {
    expect(Arr.compareForSort(10, "9")).toBeLessThan(0);
  });
});

describe("Arr.sortAscending", () => {
  it("returns an ascending copy and leaves its input alone", () => {
    const nums = [3, 1, 2];
    expect(Arr.sortAscending(nums)).toEqual([1, 2, 3]);
    expect(nums).toEqual([3, 1, 2]);
  });

  it("works when passed detached from the bundle", () => {
    const { sortAscending } = Arr;
    expect(sortAscending(["b", "a"])).toEqual(["a", "b"]);
  });
});

describe("Arr.sortDescending", () => {
  it("returns a descending copy and leaves its input alone", () => {
    const nums = [1, 3, 2];
    expect(Arr.sortDescending(nums)).toEqual([3, 2, 1]);
    expect(nums).toEqual([1, 3, 2]);
  });
});

describe("Arr.oneOrThrow", () => {
  it("returns the only item", () => {
    expect(Arr.oneOrThrow(["a"])).toBe("a");
  });

  it("throws naming the empty case for an empty array", () => {
    expect(() => Arr.oneOrThrow([])).toThrowError("This array is empty.");
  });

  it("throws naming the extra items for more than one", () => {
    expect(() => Arr.oneOrThrow(["a", "b"])).toThrowError(
      "There is more than one item in this array.",
    );
  });
});

describe("Arr.firstOrThrow", () => {
  it("returns the first item, or throws for an empty array", () => {
    expect(Arr.firstOrThrow(["a", "b"])).toBe("a");
    expect(() => Arr.firstOrThrow([])).toThrowError("This array is empty.");
  });
});

describe("Arr.lastOrThrow", () => {
  it("returns the last item, or throws for an empty array", () => {
    expect(Arr.lastOrThrow(["a", "b"])).toBe("b");
    expect(() => Arr.lastOrThrow([])).toThrowError(
      "This array has no last value—it has no value.",
    );
  });
});

describe("Arr.getOnlyItem", () => {
  it("returns the only item", () => {
    expect(Arr.getOnlyItem(["a"])).toBe("a");
  });

  it("throws naming what the array holds when it is empty or has too many", () => {
    expect(() => Arr.getOnlyItem([])).toThrowError(
      "The array does not have any items",
    );
    expect(() => Arr.getOnlyItem(["a", "b"], "rows")).toThrowError(
      "The array has too many rows",
    );
  });
});

describe("Arr.nextRotatingValue", () => {
  it("returns the item after the current one, wrapping to the start", () => {
    expect(Arr.nextRotatingValue(["a", "b", "c"], "a")).toBe("b");
    expect(Arr.nextRotatingValue(["a", "b", "c"], "c")).toBe("a");
  });

  it("returns the first item when the current one is absent", () => {
    expect(Arr.nextRotatingValue(["a", "b"], "z")).toBe("a");
  });

  it("throws for an empty array", () => {
    expect(() => Arr.nextRotatingValue([], "a")).toThrowError(
      "Cannot get next rotating value of an empty array.",
    );
  });
});

describe("Arr.insert", () => {
  it("returns a copy with the value inserted at the index", () => {
    const letters = ["a", "c"];
    expect(Arr.insert(letters, "b", 1)).toEqual(["a", "b", "c"]);
    expect(letters).toEqual(["a", "c"]);
  });
});

describe("Arr.replaceAtIndex", () => {
  it("returns a copy with the index's value replaced", () => {
    const letters = ["a", "b"];
    expect(Arr.replaceAtIndex(letters, "z", 1)).toEqual(["a", "z"]);
    expect(letters).toEqual(["a", "b"]);
  });
});

describe("Arr.replaceValue", () => {
  it("returns a copy with every match replaced", () => {
    const letters = ["a", "b", "a"];
    expect(Arr.replaceValue(letters, "a", "z")).toEqual(["z", "b", "z"]);
    expect(letters).toEqual(["a", "b", "a"]);
  });

  it("returns an equal copy when nothing matches", () => {
    expect(Arr.replaceValue(["a"], "q", "z")).toEqual(["a"]);
  });
});

describe("Arr.upOneDimension", () => {
  it("chunks the items into rows of the given length", () => {
    expect(Arr.upOneDimension([1, 2, 3, 4, 5], 2)).toEqual([
      [1, 2],
      [3, 4],
      [5],
    ]);
    expect(Arr.upOneDimension([], 2)).toEqual([[]]);
  });
});

describe("Arr.removeFirstMatchOrThrow", () => {
  it("returns a copy without the first match", () => {
    const letters = ["a", "b", "a"];
    expect(Arr.removeFirstMatchOrThrow(letters, "a")).toEqual(["b", "a"]);
    expect(letters).toEqual(["a", "b", "a"]);
  });

  it("throws when nothing matches", () => {
    expect(() => Arr.removeFirstMatchOrThrow(["a"], "z")).toThrowError(
      'No value in the array matches "z".',
    );
  });
});

describe("Arr.removeFirstMatchInPlace", () => {
  it("removes the first match from the array it was given", () => {
    const letters = ["a", "b", "a"];
    Arr.removeFirstMatchInPlace(letters, "a");
    expect(letters).toEqual(["b", "a"]);
  });
});

describe("Arr.removeAtIndex", () => {
  it("returns a copy without the index's item", () => {
    const letters = ["a", "b", "c"];
    expect(Arr.removeAtIndex(letters, 1)).toEqual(["a", "c"]);
    expect(letters).toEqual(["a", "b", "c"]);
  });

  it("throws for an index past the end", () => {
    expect(() => Arr.removeAtIndex(["a"], 1)).toThrowError(
      "The passed array does not have a value at passed idx 1",
    );
  });
});

describe("Arr.findAndRemoveFirst", () => {
  it("returns a copy without the first match and leaves its input alone", () => {
    const nums = [1, 2, 3, 2];
    expect(Arr.findAndRemoveFirst(nums, (n) => n === 2)).toEqual([1, 3, 2]);
    expect(nums).toEqual([1, 2, 3, 2]);
  });

  it("returns an equal copy when nothing matches and finding isn't required", () => {
    expect(Arr.findAndRemoveFirst([1], (n) => n === 2)).toEqual([1]);
  });

  it("throws when nothing matches and finding is required", () => {
    expect(() => Arr.findAndRemoveFirst([1], (n) => n === 2, true)).toThrowError(
      "Value not found to remove.",
    );
  });
});

describe("Arr.removeLast", () => {
  it("returns a copy without the last item", () => {
    const letters = ["a", "b"];
    expect(Arr.removeLast(letters)).toEqual(["a"]);
    expect(letters).toEqual(["a", "b"]);
    expect(Arr.removeLast([])).toEqual([]);
  });
});

describe("Arr.hasDuplicates", () => {
  it("says whether any value repeats", () => {
    expect(Arr.hasDuplicates([1, 2, 1])).toBe(true);
    expect(Arr.hasDuplicates([1, 2])).toBe(false);
  });
});

describe("Arr.includes", () => {
  it("says whether the element is in the array", () => {
    const letters = ["a", "b"] as const;
    expect(Arr.includes(letters, "a")).toBe(true);
    expect(Arr.includes(letters, "z")).toBe(false);
  });
});

describe("Arr.has", () => {
  it("says whether any item passes the test", () => {
    expect(Arr.has([1, 2], (n) => n === 2)).toBe(true);
    expect(Arr.has([1, 2], (n) => n === 3)).toBe(false);
  });
});

describe("Arr.findAll", () => {
  it("returns every item the test accepts, in order", () => {
    const nums = [1, 2, 3, 4];
    expect(Arr.findAll(nums, (n) => n % 2 === 0)).toEqual([2, 4]);
    expect(nums).toEqual([1, 2, 3, 4]);
  });
});

describe("Arr.exclude", () => {
  it("keeps the items of the first array not in the second", () => {
    expect(Arr.exclude(["a", "b", "c"], ["b", "z"])).toEqual(["a", "c"]);
  });
});

describe("Arr.excludeStrict", () => {
  it("keeps the items not among the rest arguments", () => {
    expect(Arr.excludeStrict(["a", "b", "c"], "a", "c")).toEqual(["b"]);
  });
});

describe("Arr.extract", () => {
  it("keeps the items of the first array also in the second, in the first's order", () => {
    expect(Arr.extract(["a", "b", "c"], ["c", "a", "z"])).toEqual(["a", "c"]);
  });
});

describe("Arr.extractStrict", () => {
  it("keeps the items among the rest arguments, in the array's order", () => {
    expect(Arr.extractStrict(["a", "b", "c"], "c", "a")).toEqual(["a", "c"]);
  });
});

describe("Arr.extractOrder", () => {
  it("keeps the items of the second array also in the first, in the second's order", () => {
    expect(Arr.extractOrder(["a", "b", "c"], ["c", "a"])).toEqual(["c", "a"]);
  });
});

describe("Arr.combineWithoutIdenticals", () => {
  it("joins both arrays, keeping each value's first appearance", () => {
    expect(Arr.combineWithoutIdenticals([1, 2], [2, 3, 1])).toEqual([1, 2, 3]);
  });
});
