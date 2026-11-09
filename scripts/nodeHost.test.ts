import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { describe, expect, it } from "vitest";

import { claspCredential } from "./nodeHost.ts";

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
