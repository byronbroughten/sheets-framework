import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, rmSync, symlinkSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

import { afterAll, describe, expect, it } from "vitest";

const binPath = fileURLToPath(
  new URL("./sheets-framework.ts", import.meta.url),
);
const installDir = mkdtempSync(join(tmpdir(), "sheets-framework-bin-"));

afterAll(() => {
  rmSync(installDir, { recursive: true, force: true });
});

// npm installs the bin as a symlink in node_modules/.bin, where Node refuses to strip types unless it follows the link out.
describe("the sheets-framework bin", () => {
  it("prints the CLI's usage when run through a node_modules/.bin symlink", () => {
    const binDir = join(installDir, "node_modules", ".bin");
    mkdirSync(binDir, { recursive: true });
    const linkPath = join(binDir, "sheets-framework");
    symlinkSync(binPath, linkPath);

    const { status, stderr } = spawnSync(process.execPath, [linkPath], {
      cwd: installDir,
      encoding: "utf8",
    });

    expect(stderr).toContain("Usage: sheets-framework <command> [args]");
    expect(status).toBe(1);
  });
});
