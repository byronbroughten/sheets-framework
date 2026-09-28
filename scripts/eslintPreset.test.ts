import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { eslintPreset } from "@byronbroughten/config/eslint";
import { ESLint } from "eslint";
import { defineConfig } from "eslint/config";
import { describe, expect, it } from "vitest";

import { appEslintPreset } from "./eslintPreset.js";

const appDir = mkdtempSync(join(tmpdir(), "eslint-preset-"));
const eslint = new ESLint({
  cwd: appDir,
  overrideConfigFile: true,
  overrideConfig: defineConfig(
    ...eslintPreset,
    ...appEslintPreset({ testSetupFiles: [] }),
  ),
});

async function restrictedImports(
  filePath: string,
  importPath: string,
): Promise<string[]> {
  const code = `import { x } from "${importPath}";\n\nexport const y = x;\n`;
  const results = await eslint.lintText(code, {
    filePath: join(appDir, filePath),
  });
  return results
    .flatMap(({ messages }) => messages)
    .filter(({ ruleId }) => ruleId === "no-restricted-imports")
    .map(({ message }) => message);
}

describe("appEslintPreset's generated/ boundary", () => {
  it("reports a config file imported from outside generated/", async () => {
    expect(
      await restrictedImports("src/deep/app.ts", "../generated/columnConfigs"),
    ).toHaveLength(1);
  });

  it("allows generated/appConfigs from outside generated/", async () => {
    expect(
      await restrictedImports("src/app.test.ts", "./generated/appConfigs"),
    ).toEqual([]);
  });

  it("lets a file inside generated/ import its siblings", async () => {
    expect(
      await restrictedImports("src/generated/appConfigs.ts", "./columnConfigs"),
    ).toEqual([]);
  });
});
