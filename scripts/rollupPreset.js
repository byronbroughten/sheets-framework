// @ts-check
// JS, not TypeScript: a consumer's rollup config imports this from node_modules, where Node won't strip types.
import { resolve } from "node:path";

import typescript from "@rollup/plugin-typescript";
import { parseAst } from "rollup/parseAst";

/**
 * @param {string} code
 * @returns {string[]}
 */
export function entryFunctionNames(code) {
  return parseAst(code)
    .body.filter(
      /** @returns {node is import("estree").FunctionDeclaration} */
      (node) => node.type === "FunctionDeclaration",
    )
    .map((node) => node.id.name);
}

/**
 * @param {string} code
 * @returns {string}
 */
export function exportEntryFunctions(code) {
  const names = entryFunctionNames(code);
  if (names.length === 0) return code;
  return `${code}\nexport { ${names.join(", ")} };\n`;
}

/**
 * @param {string} code
 * @returns {string}
 */
export function stripExportStatements(code) {
  const exportStatements = parseAst(code).body.filter(
    /** @returns {node is import("estree").ExportNamedDeclaration & { start: number, end: number }} */
    (node) => node.type === "ExportNamedDeclaration" && !node.source,
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
/**
 * @param {import("estree").Identifier | import("estree").Literal} node
 * @returns {string}
 */
function nameOf(node) {
  return /** @type {import("estree").Identifier} */ (node).name;
}

/** @returns {import("rollup").Plugin} */
function keepEntryFunctionsAsGlobals() {
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

/** @type {import("rollup").WarningHandlerWithDefault} */
function failOnUnresolvedImport(warning, defaultHandler) {
  if (warning.code === "UNRESOLVED_IMPORT") throw new Error(warning.message);
  defaultHandler(warning);
}

/**
 * @typedef {object} RollupPresetOptions
 * @property {string} input
 * @property {boolean} [treeshake]
 * @property {string} [tsconfig]
 * @property {string} [rootDir]
 */

// The TypeScript plugin treats a file outside rootDir as external, so rootDir widens to cover every package that gets bundled.
/**
 * @param {RollupPresetOptions} options
 * @returns {import("rollup").RollupOptions & { output: import("rollup").OutputOptions }}
 */
export function rollupPreset({
  input,
  treeshake = true,
  tsconfig = "./tsconfig.json",
  rootDir = ".",
}) {
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
          declaration: false,
          declarationMap: false,
        },
      }),
      keepEntryFunctionsAsGlobals(),
    ],
  };
}
