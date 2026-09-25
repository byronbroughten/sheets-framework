import { describe, expect, it, vi } from "vitest";

import { lazy } from "./lazy";

describe("lazy", () => {
  it("doesn't compute until first read", () => {
    const make = vi.fn(() => 1);
    lazy(make);
    expect(make).not.toHaveBeenCalled();
  });

  it("computes once and hands back the same value on every read", () => {
    const make = vi.fn(() => ({ n: 1 }));
    const read = lazy(make);
    const first = read();
    expect(read()).toBe(first);
    expect(make).toHaveBeenCalledTimes(1);
  });
});
