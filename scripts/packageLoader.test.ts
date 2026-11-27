import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { describe, expect, it } from "vitest";

import { loadDropdownStyleReset } from "./packageLoader.ts";
import type { SheetsConfig } from "./sheetsConfig.ts";

// A sheets config whose endpoint module is a temp file holding `source`, or none when `source` is undefined.
function configWithEndpointModule(source: string | undefined): SheetsConfig {
  const dir = mkdtempSync(join(tmpdir(), "endpoint-module-"));
  let endpointModule: string | undefined;
  if (source !== undefined) {
    endpointModule = join(dir, "endpoints.mjs");
    writeFileSync(endpointModule, source);
  }
  return {
    path: join(dir, "sheets.config.json"),
    dir,
    spreadsheetId: "spreadsheet-under-test",
    generatedDir: dir,
    choreHomes: [],
    endpointModule,
  };
}

describe("loadDropdownStyleReset", () => {
  it("is undefined when the app has no endpoint module", async () => {
    expect(
      await loadDropdownStyleReset(configWithEndpointModule(undefined)),
    ).toBeUndefined();
  });

  it("is undefined when the endpoint module exports none", async () => {
    const config = configWithEndpointModule("export const endpoints = {};");
    expect(await loadDropdownStyleReset(config)).toBeUndefined();
  });

  it("reads the endpoint module's export", async () => {
    const config = configWithEndpointModule(
      'export const dropdownStyleReset = "allow";',
    );
    expect(await loadDropdownStyleReset(config)).toBe("allow");
  });

  it("rejects any other value, naming the module", async () => {
    const config = configWithEndpointModule(
      'export const dropdownStyleReset = "yes";',
    );
    await expect(loadDropdownStyleReset(config)).rejects.toThrow(
      /endpoints\.mjs exports a dropdownStyleReset of "yes"/,
    );
  });
});
