import { afterAll, describe, expect, it } from "vitest";

import { appRuleMessages, removeAppDir } from "./eslintPresetFixture.ts";

async function restrictedImports(
  filePath: string,
  importPath: string,
): Promise<string[]> {
  return appRuleMessages({
    ruleId: "no-restricted-imports",
    filePath,
    code: `import { x } from "${importPath}";\n\nexport const y = x;\n`,
  });
}

describe("appEslintPreset's generated/ boundary", () => {
  afterAll(removeAppDir);

  it("reports a config file imported from outside generated/", async () => {
    expect(
      await restrictedImports(
        "src/deep/app.ts",
        "../generated/columnConfigs.ts",
      ),
    ).toEqual([expect.stringContaining("import only generated/appConfigs")]);
  });

  it("allows generated/appConfigs from outside generated/", async () => {
    expect(
      await restrictedImports("src/index.ts", "./generated/appConfigs.ts"),
    ).toEqual([]);
    expect(
      await restrictedImports("src/app.test.ts", "./generated/appConfigs.ts"),
    ).toEqual([]);
  });

  it("lets a file inside generated/ import its siblings", async () => {
    expect(
      await restrictedImports(
        "src/generated/appConfigs.ts",
        "./columnConfigs.ts",
      ),
    ).toEqual([]);
  });
});
