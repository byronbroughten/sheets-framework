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

  it("refuses a config with a missing or blank spreadsheet ID", () => {
    const { spreadsheetId: _, ...noId } = app;
    expect(() => loadSheetsConfig(repoWith({ ".": noId }))).toThrow(
      /spreadsheetId/,
    );
    const blankId = repoWith({ ".": { ...app, spreadsheetId: "" } });
    expect(() => loadSheetsConfig(blankId)).toThrow(/spreadsheetId/);
  });
});

describe("loadSheetsConfig, the folder defaults", () => {
  it("resolves the conventional folders against the config's folder when only the spreadsheet ID is set", () => {
    const root = repoWith({ app: { spreadsheetId: "app-id" } });
    mkdirSync(join(root, "app", "src"), { recursive: true });
    const config = loadSheetsConfig(join(root, "app", "src"));
    expect(config.generatedDir).toBe(join(root, "app", "src", "generated"));
    expect(config.choreHomes).toEqual([
      join(root, "app", "src", "chores"),
      join(root, "app", "src", "chores", "oneOff"),
    ]);
  });

  it("takes each explicit key in place of its default, never merged with it", () => {
    const root = repoWith({ ".": dev });
    const config = loadSheetsConfig(root);
    expect(config.generatedDir).toBe(join(root, "generated"));
    expect(config.choreHomes).toEqual([join(root, "chores")]);
  });

  it("resolves an explicit empty choreHomes to no homes", () => {
    const root = repoWith({ ".": { spreadsheetId: "app-id", choreHomes: [] } });
    expect(loadSheetsConfig(root).choreHomes).toEqual([]);
  });

  it("refuses a blank generatedDir, naming it", () => {
    const root = repoWith({ ".": { ...app, generatedDir: "" } });
    expect(() => loadSheetsConfig(root)).toThrow(/generatedDir/);
  });

  it("refuses a choreHomes that isn't an array of folder strings, naming it", () => {
    const notArray = repoWith({ ".": { ...app, choreHomes: "src/chores" } });
    expect(() => loadSheetsConfig(notArray)).toThrow(/choreHomes/);
    const blankHome = repoWith({ ".": { ...app, choreHomes: [""] } });
    expect(() => loadSheetsConfig(blankHome)).toThrow(/choreHomes/);
  });
});

describe("loadSheetsConfig, the endpoint module", () => {
  function withFile(root: string, path: string): string {
    mkdirSync(join(root, path, ".."), { recursive: true });
    writeFileSync(join(root, path), "");
    return join(root, path);
  }

  it("finds the endpoint module at the conventional path with no key", () => {
    const root = repoWith({ ".": app });
    const path = withFile(root, "src/businessEndpoints.ts");
    expect(loadSheetsConfig(root).endpointModule).toBe(path);
  });

  it("has none when the conventional path holds no file and no key is set", () => {
    const root = repoWith({ ".": app });
    expect(loadSheetsConfig(root).endpointModule).toBeUndefined();
  });

  it("takes the configured path over the conventional one", () => {
    const root = repoWith({
      ".": { ...app, endpointModule: "src/api/myEndpoints.ts" },
    });
    withFile(root, "src/businessEndpoints.ts");
    const path = withFile(root, "src/api/myEndpoints.ts");
    expect(loadSheetsConfig(root).endpointModule).toBe(path);
  });

  it("refuses a configured path that holds no file, naming it", () => {
    const root = repoWith({
      ".": { ...app, endpointModule: "src/typo.ts" },
    });
    expect(() => loadSheetsConfig(root)).toThrow(/src\/typo\.ts/);
  });

  it("refuses a configured path that isn't a filled string", () => {
    const root = repoWith({ ".": { ...app, endpointModule: "" } });
    expect(() => loadSheetsConfig(root)).toThrow(/endpointModule/);
  });
});
