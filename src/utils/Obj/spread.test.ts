import { describe, it } from "vitest";

import { assertType, type IsExactly } from "../../testSupport/typeAssertions";
import type { Spread } from "./spread";

interface Sample1 {
  a: 1;
  b: 2;
}
interface Sample2 {
  a: 2;
  b: 2;
  c: 3;
}
interface Sample3 {
  a: 3;
  d: 4;
}

describe("Spread", () => {
  it("merges left to right, a later member winning a shared key", () => {
    assertType<
      IsExactly<
        Spread<[Spread<[Sample1, Sample2]>, Sample3]>,
        { a: 3; b: 2; c: 3; d: 4 }
      >
    >(true);
  });
});
