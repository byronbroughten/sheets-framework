import { mkdirSync, mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { describe, expect, it } from "vitest";

import { ChoreIndex } from "./choreIndex.ts";

function homesWith(files: string[]): string {
  const root = mkdtempSync(join(tmpdir(), "chore-index-"));
  for (const file of files) {
    mkdirSync(join(root, file, ".."), { recursive: true });
    writeFileSync(join(root, file), "");
  }
  return root;
}

describe("ChoreIndex", () => {
  it("lists the generic chores beside the package's own, skipping Chore.ts and subfolders", () => {
    const root = homesWith([
      "generic/Chore.ts",
      "generic/fillMissingRowIds.ts",
      "generic/helpers/x.ts",
      "app/oneOff/fixUnits.ts",
    ]);
    const index = ChoreIndex.init({
      genericHome: join(root, "generic"),
      packageHomes: [join(root, "app/oneOff")],
    });
    expect(index.pathOf("fillMissingRowIds")).toBe(
      join(root, "generic", "fillMissingRowIds.ts"),
    );
    expect(index.pathOf("fixUnits")).toBe(
      join(root, "app", "oneOff", "fixUnits.ts"),
    );
    expect(index.pathOf("Chore")).toBeUndefined();
    expect(index.pathOf("x")).toBeUndefined();
  });

  it("refuses a package chore with a generic chore's name", () => {
    const root = homesWith([
      "generic/fillMissingRowIds.ts",
      "app/fillMissingRowIds.ts",
    ]);
    expect(() =>
      ChoreIndex.init({
        genericHome: join(root, "generic"),
        packageHomes: [join(root, "app")],
      }),
    ).toThrow(/shadows the generic chore/);
  });

  it("refuses one name in two package homes", () => {
    const root = homesWith(["app/a/fixUnits.ts", "app/b/fixUnits.ts"]);
    expect(() =>
      ChoreIndex.init({
        genericHome: join(root, "generic"),
        packageHomes: [join(root, "app/a"), join(root, "app/b")],
      }),
    ).toThrow(/more than one chore home/);
  });
});
