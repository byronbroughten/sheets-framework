import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

import { describe, expect, it } from "vitest";

const scriptPath = fileURLToPath(
  new URL("./setup-clasp-run-auth.sh", import.meta.url),
);

// The wizard opens a browser and touches credentials, so a parse is as far as a test can run it.
describe("setup-clasp-run-auth.sh", () => {
  it("parses under bash -n", () => {
    const { status, stderr } = spawnSync("bash", ["-n", scriptPath], {
      encoding: "utf8",
    });
    expect(stderr).toBe("");
    expect(status).toBe(0);
  });
});
