import { afterAll, describe, expect, it } from "vitest";

import { appRuleMessages, removeAppDir } from "./eslintPresetFixture.ts";

async function restrictedSyntax(code: string): Promise<string[]> {
  return appRuleMessages({
    ruleId: "no-restricted-syntax",
    filePath: "src/app.ts",
    code,
  });
}

function subclassWith(field: string): string {
  return `class Base {}\n\nexport class Sub extends Base {\n  ${field}\n}\n`;
}

describe("appEslintPreset's subclass fields", () => {
  afterAll(removeAppDir);

  it("reports a subclass field with neither an initializer nor declare", async () => {
    expect(await restrictedSyntax(subclassWith("count: number;"))).toEqual([
      expect.stringContaining("declare"),
    ]);
  });

  it("allows a declare field and an initialised one", async () => {
    expect(
      await restrictedSyntax(subclassWith("declare count: number;")),
    ).toEqual([]);
    expect(await restrictedSyntax(subclassWith("count = 1;"))).toEqual([]);
  });

  it("still reports config's style selectors alongside the subclass one", async () => {
    expect(
      await restrictedSyntax("export enum Kind {\n  One,\n}\n"),
    ).not.toEqual([]);
  });
});
