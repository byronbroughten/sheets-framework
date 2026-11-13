import { describe, expect, it } from "vitest";

import { parseConfirmToken } from "./heldFile.ts";

describe("parseConfirmToken", () => {
  it("reads the token after --confirm", () => {
    expect(parseConfirmToken(["--send", "--confirm", "7f3a9c"])).toBe("7f3a9c");
  });

  it("is undefined without the flag", () => {
    expect(parseConfirmToken(["--send"])).toBeUndefined();
  });

  it("rejects --confirm with no token", () => {
    expect(() => parseConfirmToken(["--send", "--confirm"])).toThrow(
      "--confirm takes the token the dry run printed.",
    );
  });
});
