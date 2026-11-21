import { resolve } from "node:path";

import typescriptModule from "@rollup/plugin-typescript";
import type {
  ExportNamedDeclaration,
  FunctionDeclaration,
  Identifier,
  Literal,
} from "estree";
import type {
  LoggingFunction,
  OutputOptions,
  Plugin,
  RollupLog,
  RollupOptions,
} from "rollup";
import { parseAst } from "rollup/parseAst";

// Its one types file reads as CommonJS under nodenext, but Node loads its ESM build, whose default export is the plugin itself.
const typescript =
  typescriptModule as unknown as typeof typescriptModule.default;

export function entryFunctionNames(code: string): string[] {
  return parseAst(code)
    .body.filter(
      (node): node is FunctionDeclaration =>
        node.type === "FunctionDeclaration",
    )
    .map((node) => node.id.name);
}

export function exportEntryFunctions(code: string): string {
  const names = entryFunctionNames(code);
  if (names.length === 0) return code;
  return `${code}\nexport { ${names.join(", ")} };\n`;
}

export function stripExportStatements(code: string): string {
  const exportStatements = parseAst(code).body.filter(
    (node): node is ExportNamedDeclaration & { start: number; end: number } =>
      node.type === "ExportNamedDeclaration" && !node.source,
  );
  for (const { specifiers } of exportStatements) {
    for (const specifier of specifiers) {
      const local = nameOf(specifier.local);
      const exported = nameOf(specifier.exported);
      if (local !== exported) {
        throw new Error(
          `Rollup renamed an entry function to avoid a name clash (${local} as ${exported}). Apps Script calls it by its original name, so rename one of the two functions.`,
        );
      }
    }
  }
  return exportStatements.reduceRight(
    (stripped, { start, end }) =>
      stripped.slice(0, start) +
      stripped.slice(stripped[end] === "\n" ? end + 1 : end),
    code,
  );
}

// Rollup names its own exports with identifiers, never string literals.
function nameOf(node: Identifier | Literal): string {
  return (node as Identifier).name;
}

function keepEntryFunctionsAsGlobals(): Plugin {
  return {
    name: "keep-entry-functions-as-globals",
    transform(code, id) {
      if (!this.getModuleInfo(id)?.isEntry) return null;
      return { code: exportEntryFunctions(code), map: null };
    },
    renderChunk(code, chunk) {
      if (!chunk.isEntry) return null;
      return { code: stripExportStatements(code), map: null };
    },
  };
}

function failOnUnresolvedImport(
  warning: RollupLog,
  defaultHandler: LoggingFunction,
): void {
  if (warning.code === "UNRESOLVED_IMPORT") throw new Error(warning.message);
  defaultHandler(warning);
}

interface RollupPresetOptions {
  input: string;
  treeshake?: boolean;
  tsconfig?: string;
  rootDir?: string;
}

// The TypeScript plugin treats a file outside rootDir as external, so rootDir widens to cover every package that gets bundled.
export function rollupPreset({
  input,
  treeshake = true,
  tsconfig = "./tsconfig.json",
  rootDir = ".",
}: RollupPresetOptions): RollupOptions & { output: OutputOptions } {
  return {
    input,
    output: { file: "dist/bundle.js", format: "es", sourcemap: true },
    treeshake,
    onwarn: failOnUnresolvedImport,
    plugins: [
      typescript({
        tsconfig,
        filterRoot: false,
        compilerOptions: {
          rootDir: resolve(rootDir),
          // The plugin's own module default would clash with the tsconfig's nodenext resolution (TS5110).
          module: "nodenext",
          // The sources import each other by .ts; without it the plugin warns TS5096.
          rewriteRelativeImportExtensions: true,
          declaration: false,
          declarationMap: false,
        },
      }),
      keepEntryFunctionsAsGlobals(),
    ],
  };
}
