import { eslintPreset, variableNaming } from "@byronbroughten/config/eslint";
import { defineConfig } from "eslint/config";

import {
  platformImportPattern,
  sheetsSrcBlocks,
} from "./scripts/eslintPreset.js";

const importPatterns = {
  raw: {
    regex:
      "^(\\.\\./)+01_SpreadsheetSchema/(SheetSchema|ColumnSchema|columnConfigsTypes|valueConfigsTypes|generated/(columnConfigs|valueConfigs))(\\.js)?$",
    message:
      "Raw is positional: it addresses by GID and index and never resolves a column. Column and value lookups belong in the Identified tier or above.",
  },
  appConfigs: {
    regex: "(^|/)generated/|(^|/)appConfigs(\\.js)?$",
    message:
      "Tiers take config types from Register and values from installedConfigs(); only the app's appConfigs.ts and the framework's dev/devConfigs.ts import generated configs.",
  },
};
// utils/ sits below every numbered tier.
const layerFolders = [
  "utils",
  "00_Source",
  "01_SpreadsheetSchema",
  "02_SpreadsheetRaw",
  "03_SpreadsheetIdentified",
  "04_SpreadsheetNamed",
  "05_Operators",
  "06_API",
];
// The framework's un-numbered folders built on the tiers; utils/ and testSupport/ are not among them.
const aboveTierFolders = [
  "appsScriptHost",
  "chores",
  "framework",
  "frameworkTesting",
  "nodeHost",
];
function layerImportPattern(layer) {
  return {
    regex: `(^|/)(${[...layerFolders.slice(layer + 1), ...aboveTierFolders].join("|")})(/|(\\.js)?$)`,
    message: `Dependencies only point downward: ${layerFolders[layer]} imports nothing from a higher tier or from ${aboveTierFolders.join(", ")} (packages/framework/src/AGENTS.md).`,
  };
}
// After the platform block: a later block's no-restricted-imports replaces an earlier one's, so each merges the patterns that still apply.
const layerImportBlocks = layerFolders.flatMap((folder, layer) => {
  const extra = folder === "02_SpreadsheetRaw" ? [importPatterns.raw] : [];
  function restrict(patterns) {
    return {
      "no-restricted-imports": [
        "error",
        { patterns: [...patterns, layerImportPattern(layer), ...extra] },
      ],
    };
  }
  const isPlatformFolder = folder === "00_Source";
  return [
    {
      files: [`src/${folder}/**/*.ts`],
      ignores: [
        "**/*.test.ts",
        ...(isPlatformFolder ? ["src/00_Source/GoogleSheets/**"] : []),
      ],
      rules: restrict([platformImportPattern, importPatterns.appConfigs]),
    },
    {
      files: [
        `src/${folder}/**/*.test.ts`,
        ...(isPlatformFolder ? ["src/00_Source/GoogleSheets/**/*.ts"] : []),
      ],
      rules: restrict([]),
    },
  ];
});

export default defineConfig(
  ...eslintPreset,
  // tsc checks these for undefined names, as typescript-eslint leaves it to tsc in .ts files.
  { files: ["scripts/**/*.js"], rules: { "no-undef": "off" } },
  // Domain-free utilities and type assertions keep bare T, K and V.
  {
    files: ["src/utils/**/*.ts", "src/testSupport/typeAssertions.ts"],
    rules: {
      "@typescript-eslint/naming-convention": ["error", variableNaming],
    },
  },
  ...sheetsSrcBlocks([
    "src/00_Source/GoogleSheets/**",
    "src/appsScriptHost/**",
    "src/nodeHost/**",
    "src/chores/**",
    "src/testSupport/**",
    "src/TypeDeclarations/**",
  ]),
  ...layerImportBlocks,
);
