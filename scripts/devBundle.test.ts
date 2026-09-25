import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { runInNewContext } from "node:vm";

import { rollup } from "rollup";
import { describe, expect, it } from "vitest";

import { rollupPreset } from "./rollupPreset.js";

const frameworkRoot = join(dirname(fileURLToPath(import.meta.url)), "..");

async function buildDevBundle(): Promise<string> {
  const config = rollupPreset({
    input: join(frameworkRoot, "dev/index.ts"),
    tsconfig: join(frameworkRoot, "tsconfig.json"),
    rootDir: frameworkRoot,
  });
  const bundle = await rollup({ ...config, onwarn: () => {} });
  const { output } = await bundle.generate(config.output);
  return output[0].code;
}

describe("the dev bundle", () => {
  // Rollup orders a circular import by the first file that enters it, and only a real load shows a class extended before its definition.
  it("loads in a bare context and defines both triggers as globals", async () => {
    const code = await buildDevBundle();
    const context: { triggerOnEdit?: unknown; triggerOnChange?: unknown } = {};
    runInNewContext(
      `${code}\nthis.triggerOnEdit = triggerOnEdit; this.triggerOnChange = triggerOnChange;`,
      context,
    );
    expect(typeof context.triggerOnEdit).toBe("function");
    expect(typeof context.triggerOnChange).toBe("function");
  }, 60_000);
});
