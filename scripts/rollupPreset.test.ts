import { mkdirSync, mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

import { rollup } from "rollup";
import { describe, expect, it } from "vitest";

import {
  entryFunctionNames,
  exportEntryFunctions,
  rollupPreset,
  stripExportStatements,
} from "./rollupPreset.js";

describe("entryFunctionNames", () => {
  it("lists top-level function declarations in source order", () => {
    const code = `function b() {}\nasync function a() {}\nfunction* c() {}\n`;
    expect(entryFunctionNames(code)).toEqual(["b", "a", "c"]);
  });

  it("skips functions already exported", () => {
    const code = `export function a() {}\nfunction b() {}\nexport default function c() {}\n`;
    expect(entryFunctionNames(code)).toEqual(["b"]);
  });

  it("skips nested functions, arrow functions and classes", () => {
    const code = `function a() { function inner() {} }\nconst b = () => {};\nclass C { d() {} }\n`;
    expect(entryFunctionNames(code)).toEqual(["a"]);
  });

  it("ignores text that only looks like a declaration", () => {
    const code = `// function fake() {}\nconst s = "\\nfunction alsoFake() {}";\nfunction real() {}\n`;
    expect(entryFunctionNames(code)).toEqual(["real"]);
  });
});

describe("exportEntryFunctions", () => {
  it("appends one export statement naming every unexported function", () => {
    expect(exportEntryFunctions("function a() {}\nfunction b() {}\n")).toBe(
      "function a() {}\nfunction b() {}\n\nexport { a, b };\n",
    );
  });

  it("leaves code with no such function untouched", () => {
    const code = "export function a() {}\nconst b = 1;\n";
    expect(exportEntryFunctions(code)).toBe(code);
  });
});

describe("stripExportStatements", () => {
  it("removes the export list and keeps the declarations", () => {
    const chunk = "function a() {}\nfunction b() {}\n\nexport { a, b };\n";
    expect(stripExportStatements(chunk)).toBe(
      "function a() {}\nfunction b() {}\n\n",
    );
  });

  it("removes a multi-line export list", () => {
    const chunk =
      "function a() {}\nfunction b() {}\nexport {\n  a,\n  b,\n};\n";
    expect(stripExportStatements(chunk)).toBe(
      "function a() {}\nfunction b() {}\n",
    );
  });

  it("keeps a re-export from another module", () => {
    const chunk = `function a() {}\nexport { x } from "elsewhere";\nexport { a };\n`;
    expect(stripExportStatements(chunk)).toBe(
      `function a() {}\nexport { x } from "elsewhere";\n`,
    );
  });

  it("throws when rollup renamed an export, since the global would change name", () => {
    const chunk = "function a$1() {}\nexport { a$1 as a };\n";
    expect(() => stripExportStatements(chunk)).toThrow(/a\$1 as a/);
  });

  it("returns a chunk with no export statement unchanged", () => {
    const chunk = "function a() {}\n";
    expect(stripExportStatements(chunk)).toBe(chunk);
  });
});

// The preset reads its tsconfig when called, so these point at a real one.
const tsconfig = fileURLToPath(new URL("../tsconfig.json", import.meta.url));

describe("rollupPreset", () => {
  it("returns a plain config that tree-shakes by default", () => {
    const config = rollupPreset({ input: "src/index.ts", tsconfig });
    expect(config).toMatchObject({
      input: "src/index.ts",
      output: { file: "dist/bundle.js", format: "es", sourcemap: true },
      treeshake: true,
    });
    expect(Object.getPrototypeOf(config)).toBe(Object.prototype);
  });

  it("accepts treeshake: false", () => {
    const config = rollupPreset({
      input: "src/index.ts",
      tsconfig,
      treeshake: false,
    });
    expect(config.treeshake).toBe(false);
  });
});

// Two sibling folders under one root, like the two packages after the workspace move.
function twoPackages(appSource: string): string {
  const root = mkdtempSync(join(tmpdir(), "rollup-preset-"));
  mkdirSync(join(root, "framework", "src"), { recursive: true });
  mkdirSync(join(root, "app", "src"), { recursive: true });
  writeFileSync(
    join(root, "framework", "src", "lib.ts"),
    [
      "export function usedByEntry(): string { return 'used'; }",
      "export function neverCalled(): string { return 'unused'; }",
      "",
    ].join("\n"),
  );
  writeFileSync(join(root, "app", "src", "index.ts"), appSource);
  writeFileSync(
    join(root, "app", "tsconfig.json"),
    JSON.stringify({
      compilerOptions: {
        target: "ES2022",
        module: "esnext",
        moduleResolution: "bundler",
        strict: true,
        rootDir: "./src",
        outDir: "./dist",
        declaration: true,
        declarationMap: true,
        sourceMap: true,
        skipLibCheck: true,
        types: [],
      },
      include: ["src/**/*.ts"],
    }),
  );
  return root;
}

async function bundleOf(
  root: string,
  options: { treeshake?: boolean } = {},
): Promise<string> {
  const config = rollupPreset({
    input: join(root, "app", "src", "index.ts"),
    tsconfig: join(root, "app", "tsconfig.json"),
    rootDir: root,
    ...options,
  });
  const build = await rollup(config);
  const file = join(root, "app", "dist", "bundle.js");
  const { output } = await build.generate({ ...config.output, file });
  return output[0].code;
}

describe("a build with the preset", () => {
  const appSource = [
    "import { usedByEntry } from '../../framework/src/lib';",
    "function triggerOnEdit() { return usedByEntry(); }",
    "function scratch() { return 1; }",
    "",
  ].join("\n");

  it("bundles the sibling package's source and shakes out what nothing calls", async () => {
    const code = await bundleOf(twoPackages(appSource));
    expect(code).toContain("function usedByEntry()");
    expect(code).not.toContain("neverCalled");
  }, 60_000);

  it("keeps every entry function as a global, uncalled ones too, with no export statement", async () => {
    const code = await bundleOf(twoPackages(appSource));
    expect(code).toContain("function triggerOnEdit()");
    expect(code).toContain("function scratch()");
    expect(code).not.toMatch(/^export /m);
  }, 60_000);

  it("keeps an entry function that is already exported, with no export statement", async () => {
    const code = await bundleOf(
      twoPackages("export function triggerOnEdit() { return 1; }\n"),
    );
    expect(code).toContain("function triggerOnEdit()");
    expect(code).not.toMatch(/^export /m);
  }, 60_000);

  it("keeps everything with treeshake: false", async () => {
    const code = await bundleOf(twoPackages(appSource), { treeshake: false });
    expect(code).toContain("function neverCalled()");
    expect(code).toContain("function triggerOnEdit()");
    expect(code).not.toMatch(/^export /m);
  }, 60_000);

  it("fails on an import that does not resolve", async () => {
    const root = twoPackages(
      "import { nope } from 'no-such-package';\nfunction triggerOnEdit() { return nope; }\n",
    );
    await expect(bundleOf(root)).rejects.toThrow(/no-such-package/);
  }, 60_000);
});
