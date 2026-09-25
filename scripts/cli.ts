// The sheets-framework subcommands: runs one against the package whose sheets.config.json sits above cwd. See docs/how-it-runs.md.
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

import { loadSheetsConfig } from "./sheetsConfig.ts";

const usage = `Usage: sheets-framework <command> [args]
  gen-configs          regenerate the package's four config files from its config sheets
  chore [name] [--send] [--json]
                       dry-run a chore (or apply it with --send); no name lists them
  probe --fields|--filter|--path ...
                       one read-only Sheets request; no args prints its usage
  setup-auth           mint the Node host's clasp credential

The spreadsheet ID comes only from the nearest sheets.config.json above cwd.`;

const [command, ...argv] = process.argv.slice(2);

switch (command) {
  case "gen-configs": {
    const { runGenConfigs } = await import("./genConfigs.ts");
    await runGenConfigs(loadSheetsConfig());
    break;
  }
  case "chore": {
    const { runChore } = await import("./chore.ts");
    await runChore(loadSheetsConfig(), argv);
    break;
  }
  case "probe": {
    const { runProbe } = await import("./sheetsProbe.ts");
    runProbe(loadSheetsConfig(), argv);
    break;
  }
  case "setup-auth": {
    const script = fileURLToPath(
      new URL("./setup-clasp-run-auth.sh", import.meta.url),
    );
    const { status } = spawnSync("bash", [script, ...argv], {
      stdio: "inherit",
    });
    process.exit(status ?? 1);
    break;
  }
  default:
    console.error(
      command ? `Unknown command "${command}".\n\n${usage}` : usage,
    );
    process.exit(1);
}
