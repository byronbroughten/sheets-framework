import { join } from "node:path";

import { describe, expect, it } from "vitest";

import { HeldFile, parseConfirmToken } from "./HeldFile.ts";

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

describe("HeldFile", () => {
  it("keeps a chore's held list in its own file", () => {
    const heldFile = HeldFile.init({
      packageDir: "/package",
      command: "chore",
      isSending: false,
      confirmToken: undefined,
    });
    expect(heldFile.path).toBe(join("/package", ".probe", "chore-held.json"));
  });
});
