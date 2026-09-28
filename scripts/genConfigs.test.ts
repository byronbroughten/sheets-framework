import { describe, expect, it } from "vitest";

import { appConfigsText } from "./genConfigs.ts";

describe("appConfigsText", () => {
  const text = appConfigsText();

  it("builds appConfigs from the three sibling config files", () => {
    expect(text).toContain('import { columnConfigs } from "./columnConfigs";');
    expect(text).toContain('import { sheetConfigs } from "./sheetConfigs";');
    expect(text).toContain('import { valueConfigs } from "./valueConfigs";');
    expect(text).toContain(
      "export const appConfigs = { sheetConfigs, columnConfigs, valueConfigs };",
    );
  });

  it("brings the framework's entry into the program with a type-only import", () => {
    expect(text).toContain(
      'import type {} from "@byronbroughten/sheets-framework";',
    );
  });

  it("augments the package's Register with the appConfigs type", () => {
    expect(text).toMatch(
      /declare module "@byronbroughten\/sheets-framework" \{\s*interface Register \{\s*configs: typeof appConfigs;\s*\}\s*\}/,
    );
  });
});
