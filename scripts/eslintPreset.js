// @ts-check
// JS, not TypeScript: a consumer's eslint config imports this from node_modules, where Node won't strip types.
import { styleSyntax, variableNaming } from "@byronbroughten/config/eslint";

const frameworkPackage = "@byronbroughten/sheets-framework";
const platformMessage =
  "Google Sheets code lives only in the framework's src/00_Source/GoogleSheets/. Everything else takes a platform-neutral type or a return value that the entry point handles.";
export const platformImportPattern = {
  regex: "GoogleSheets/|GoogleSheets/(GoogleSheetsAPI|AppsScript)(\\.js)?$",
  message: platformMessage,
};

// One copy for the framework and its apps, so the platform boundary can't drift between them.
/**
 * @param {string[]} platformIgnores
 * @returns {import("eslint").Linter.Config[]}
 */
export function sheetsSrcBlocks(platformIgnores) {
  return [
    // Tests keep bare T, K and V, like the domain-free utilities each package names itself.
    {
      files: ["**/*.test.ts"],
      rules: {
        "@typescript-eslint/naming-convention": ["error", variableNaming],
      },
    },
    // The structural utilities do the generic typing that needs `any` (docs/style/type-modeling.md).
    {
      files: ["src/**/{Obj,Arr}.ts", "src/**/{Obj,Arr}/**/*.ts"],
      rules: { "@typescript-eslint/no-explicit-any": "off" },
    },
    {
      files: ["src/**/*.ts"],
      ignores: [...platformIgnores, "src/index.ts", "**/*.test.ts"],
      rules: {
        "no-restricted-imports": [
          "error",
          { patterns: [platformImportPattern] },
        ],
        "no-restricted-syntax": [
          "error",
          ...styleSyntax,
          {
            selector:
              "TSQualifiedName[left.type='Identifier'][left.name='GoogleAppsScript']",
            message: platformMessage,
          },
        ],
        "no-restricted-globals": [
          "error",
          ...["SpreadsheetApp", "ScriptApp", "PropertiesService", "Sheets"].map(
            (name) => ({ name, message: platformMessage }),
          ),
        ],
      },
    },
  ];
}

// App code reaches the framework only through its package name, and a relative import never leaves the app's src/.
const appEntryMessage = `App code imports the framework only from "${frameworkPackage}", and "${frameworkPackage}/testing" only from *.test.ts (#140).`;
/**
 * @param {number} depth
 * @param {boolean} isTestingAllowed
 */
function appImportPatterns(depth, isTestingAllowed) {
  return [
    {
      regex: `^(\\.\\./){${depth + 1}}`,
      message: "A relative import stays inside the app's src/ (#140).",
    },
    {
      regex: `^${frameworkPackage}/${isTestingAllowed ? "(?!testing$)" : ""}`,
      message: appEntryMessage,
    },
  ];
}

/**
 * @typedef {object} AppEslintPresetOptions
 * @property {string[]} testSetupFiles
 */

// Layered on @byronbroughten/config's eslintPreset, which the app spreads first.
/**
 * @param {AppEslintPresetOptions} options
 * @returns {import("eslint").Linter.Config[]}
 */
export function appEslintPreset({ testSetupFiles }) {
  const appImportBlocks = [0, 1, 2, 3, 4, 5].flatMap((depth) => {
    const files = [`src/${"*/".repeat(depth)}*.ts`];
    /**
     * @param {boolean} isTestingAllowed
     * @returns {import("eslint").Linter.RulesRecord}
     */
    function restrict(isTestingAllowed) {
      return {
        "no-restricted-imports": [
          "error",
          { patterns: appImportPatterns(depth, isTestingAllowed) },
        ],
      };
    }
    return [
      {
        files,
        // generated/ imports the framework's makeConfigs by the relative path gen:configs writes.
        ignores: ["**/*.test.ts", "src/generated/**", ...testSetupFiles],
        rules: restrict(false),
      },
      {
        files: files.flatMap((glob) => [
          [glob, "**/*.test.ts"],
          ...(depth === 0 ? testSetupFiles : []),
        ]),
        rules: restrict(true),
      },
    ];
  });
  return [...sheetsSrcBlocks([]), ...appImportBlocks];
}
