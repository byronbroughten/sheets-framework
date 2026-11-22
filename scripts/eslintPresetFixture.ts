import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { eslintPreset } from "@byronbroughten/code-style/eslint";
import { ESLint, type Linter } from "eslint";
import { defineConfig } from "eslint/config";

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

interface AppLint {
  ruleId: string;
  filePath: string;
  code: string;
}

// The messages one rule reports on code linted as an app file under appEslintPreset.
export async function appRuleMessages({
  ruleId,
  filePath,
  code,
}: AppLint): Promise<string[]> {
  const results = await eslint.lintText(code, {
    filePath: join(appDir, filePath),
  });
  return results
    .flatMap(({ messages }) => messages)
    .filter((message) => message.ruleId === ruleId)
    .map(({ message }) => message);
}

export function removeAppDir(): void {
  rmSync(appDir, { recursive: true });
}
