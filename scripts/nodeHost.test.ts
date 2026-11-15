import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { describe, expect, it } from "vitest";

import { claspCredential, loadDropdownStyleReset } from "./nodeHost.ts";
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

function credentialFile(
  dir: string,
  name: string,
  user: string,
  token: string,
) {
  const file = join(dir, name);
  const credential = {
    client_id: `${token}-id`,
    client_secret: `${token}-secret`,
    refresh_token: `${token}-refresh`,
  };
  writeFileSync(file, JSON.stringify({ tokens: { [user]: credential } }));
  return file;
}

describe("claspCredential", () => {
  it("reads the desktop-clasp-run login from ~/.clasprc.json when clasp_config_auth is unset", () => {
    const home = mkdtempSync(join(tmpdir(), "clasp-home-"));
    credentialFile(home, ".clasprc.json", "desktop-clasp-run", "developer");
    expect(claspCredential({}, home).refresh_token).toBe("developer-refresh");
  });

  it("reads the file clasp_config_auth names instead when it is set", () => {
    const home = mkdtempSync(join(tmpdir(), "clasp-home-"));
    credentialFile(home, ".clasprc.json", "desktop-clasp-run", "developer");
    const agentFile = credentialFile(home, "agent.json", "default", "agent");
    expect(
      claspCredential({ clasp_config_auth: agentFile }, home).refresh_token,
    ).toBe("agent-refresh");
  });

  it("does not fall back to the developer's login when the agent file lacks a credential", () => {
    const home = mkdtempSync(join(tmpdir(), "clasp-home-"));
    credentialFile(home, ".clasprc.json", "desktop-clasp-run", "developer");
    const agentFile = credentialFile(home, "agent.json", "other", "agent");
    expect(() =>
      claspCredential({ clasp_config_auth: agentFile }, home),
    ).toThrow(`No "default" credential in ${agentFile}`);
  });
});

describe("loadDropdownStyleReset", () => {
  it('is "ask" when the app has no endpoint module', async () => {
    expect(
      await loadDropdownStyleReset(configWithEndpointModule(undefined)),
    ).toBe("ask");
  });

  it('is "ask" when the endpoint module exports none', async () => {
    const config = configWithEndpointModule("export const endpoints = {};");
    expect(await loadDropdownStyleReset(config)).toBe("ask");
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
