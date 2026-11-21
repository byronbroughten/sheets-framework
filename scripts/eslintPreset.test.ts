import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { eslintPreset } from "@byronbroughten/config/eslint";
import { ESLint, type Linter } from "eslint";
import { defineConfig } from "eslint/config";
import { afterAll, describe, expect, it } from "vitest";

import { appEslintPreset } from "./eslintPreset.ts";

const appDir = mkdtempSync(join(tmpdir(), "eslint-preset-"));
const eslint = new ESLint({
  cwd: appDir,
  overrideConfigFile: true,
  overrideConfig: defineConfig(
    // Its JS literals infer rule levels as string, which Linter's types reject.
    ...(eslintPreset as Linter.Config[]),
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

afterAll(() => {
  rmSync(appDir, { recursive: true });
});

async function restrictedSyntax(code: string): Promise<string[]> {
  const results = await eslint.lintText(code, {
    filePath: join(appDir, "src/app.ts"),
  });
  return results
    .flatMap(({ messages }) => messages)
    .filter(({ ruleId }) => ruleId === "no-restricted-syntax")
    .map(({ message }) => message);
}

function subclassWith(field: string): string {
  return `class Base {}\n\nexport class Sub extends Base {\n  ${field}\n}\n`;
}

describe("appEslintPreset's subclass fields", () => {
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

  it("keeps the style selectors config exports", async () => {
    expect(
      await restrictedSyntax("export enum Kind {\n  One,\n}\n"),
    ).not.toEqual([]);
  });
});

describe("appEslintPreset's generated/ boundary", () => {
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
