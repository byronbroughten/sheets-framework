import { describe, expect, it, vi } from "vitest";

describe("installEndpoints", () => {
  it("throws when the configs aren't installed yet", async () => {
    vi.resetModules();
    const { installEndpoints } = await import("./feedbackColumnIds");
    expect(() => installEndpoints({})).toThrow(
      "Configs have not been installed.",
    );
  });
});
