import { styleSyntax, variableNaming } from "@byronbroughten/config/eslint";
import type { Linter } from "eslint";

const requestRecordingMessage =
  "Assert the grid a write leaves through stubSheetsService()'s grid, not requests a hand-rolled batchUpdate records; only GoogleSheetsAPI.test.ts tests the request mapping (docs/code-style.md, Tests).";
const requestRecordingSyntax = [
  "ObjectExpression > Property[key.name='batchUpdate']",
  "ObjectExpression > Property[key.value='batchUpdate']",
].map((selector) => ({ selector, message: requestRecordingMessage }));

const frameworkPackage = "@byronbroughten/sheets-framework";
const platformMessage =
  "Google Sheets code lives only in the framework's src/00_Source/GoogleSheets/. Everything else takes a platform-neutral type or a return value that the entry point handles.";
interface ImportPattern {
  regex: string;
  message: string;
}

export const platformImportPattern: ImportPattern = {
  regex: "GoogleSheets/",
  message: platformMessage,
};

// One copy for the framework and its apps, so the platform boundary can't drift between them.
export function sheetsSrcBlocks(platformIgnores: string[]): Linter.Config[] {
  return [
    // Tests keep bare T, K and V, like the domain-free utilities each package names itself.
    {
      files: ["**/*.test.ts"],
      rules: {
        "@typescript-eslint/naming-convention": ["error", variableNaming],
      },
    },
    {
      files: ["**/*.test.ts"],
      ignores: ["**/GoogleSheetsAPI.test.ts"],
      rules: {
        "no-restricted-syntax": [
          "error",
          ...styleSyntax,
          ...requestRecordingSyntax,
        ],
      },
    },
    // The structural utilities do the generic typing that needs `any` (docs/code-style/type-modeling.md).
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
          {
            selector:
              ":matches(ClassDeclaration, ClassExpression)[superClass] > ClassBody > PropertyDefinition[value=null][declare!=true][static!=true]",
            message:
              "A subclass field takes an initializer or declare: under Node's type stripping a bare one resets to undefined after super(), while the Apps Script bundle keeps its value.",
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
function appImportPatterns(
  depth: number,
  isTestingAllowed: boolean,
): ImportPattern[] {
  return [
    {
      regex: `^(\\.\\./){${depth + 1}}`,
      message: "A relative import stays inside the app's src/ (#140).",
    },
    {
      regex: `^${frameworkPackage}/${isTestingAllowed ? "(?!testing$)" : ""}`,
      message: appEntryMessage,
    },
    {
      regex: "(^|/)generated/(?!appConfigs(\\.ts)?$)",
      message:
        "Outside generated/, import only generated/appConfigs; the other generated files are its internals.",
    },
  ];
}

interface AppEslintPresetOptions {
  testSetupFiles: string[];
}

// Layered on @byronbroughten/config's eslintPreset, which the app spreads first.
export function appEslintPreset({
  testSetupFiles,
}: AppEslintPresetOptions): Linter.Config[] {
  const appImportBlocks = [0, 1, 2, 3, 4, 5].flatMap((depth) => {
    const files = [`src/${"*/".repeat(depth)}*.ts`];
    function restrict(isTestingAllowed: boolean): Linter.RulesRecord {
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
        ignores: ["**/*.test.ts", ...testSetupFiles],
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
