import { mkdirSync, mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { describe, expect, it } from "vitest";

import { loadSheetsConfig } from "./sheetsConfig.ts";

function repoWith(configs: Record<string, object>): string {
  const root = mkdtempSync(join(tmpdir(), "sheets-config-"));
  mkdirSync(join(root, ".git"));
  for (const [dir, config] of Object.entries(configs)) {
    mkdirSync(join(root, dir), { recursive: true });
    writeFileSync(
      join(root, dir, "sheets.config.json"),
      JSON.stringify(config),
    );
  }
  return root;
}

const app = {
  spreadsheetId: "app-id",
  generatedDir: "src/generated",
  choreHomes: ["src/chores", "src/chores/oneOff"],
};
const dev = {
  spreadsheetId: "dev-id",
  generatedDir: "generated",
  choreHomes: ["chores"],
};

describe("loadSheetsConfig", () => {
  it("finds the nearest config by walking up from cwd and resolves its folders against it", () => {
    const root = repoWith({ ".": app, dev });
    mkdirSync(join(root, "dev", "generated"), { recursive: true });
    expect(loadSheetsConfig(join(root, "dev", "generated"))).toEqual({
      path: join(root, "dev", "sheets.config.json"),
      dir: join(root, "dev"),
      spreadsheetId: "dev-id",
      generatedDir: join(root, "dev", "generated"),
      choreHomes: [join(root, "dev", "chores")],
    });
    expect(loadSheetsConfig(join(root, "src")).spreadsheetId).toBe("app-id");
  });

  it("refuses to run without a config above cwd", () => {
    const root = repoWith({});
    expect(() => loadSheetsConfig(root)).toThrow(/No sheets\.config\.json/);
  });

  it("names the example file when a package has only that", () => {
    const root = repoWith({});
    writeFileSync(
      join(root, "sheets.config.example.json"),
      JSON.stringify(app),
    );
    expect(() => loadSheetsConfig(root)).toThrow(
      /sheets\.config\.example\.json/,
    );
    expect(() => loadSheetsConfig(root)).not.toThrow(/No sheets\.config\.json/);
  });

  it("prefers the real config over an example beside it", () => {
    const root = repoWith({ ".": app });
    writeFileSync(
      join(root, "sheets.config.example.json"),
      JSON.stringify(dev),
    );
    expect(loadSheetsConfig(root).spreadsheetId).toBe("app-id");
  });

  it("refuses two packages that share a spreadsheet ID, from either package", () => {
    const root = repoWith({
      ".": app,
      "packages/dev": { ...dev, spreadsheetId: "app-id" },
    });
    expect(() => loadSheetsConfig(root)).toThrow(/share a spreadsheet ID/);
    expect(() => loadSheetsConfig(join(root, "packages", "dev"))).toThrow(
      /share a spreadsheet ID/,
    );
  });

  it("refuses a config missing a field", () => {
    const root = repoWith({
      ".": { spreadsheetId: "app-id", generatedDir: "src/generated" },
    });
    expect(() => loadSheetsConfig(root)).toThrow(/choreHomes/);
    const blankId = repoWith({ ".": { ...app, spreadsheetId: "" } });
    expect(() => loadSheetsConfig(blankId)).toThrow(/spreadsheetId/);
  });
});
