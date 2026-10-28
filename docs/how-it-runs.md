# How it runs

**The framework runs in two hosts, and only one of them is production**: Apps Script, pushed with `clasp`, and the Node host behind the chore runner, `gen-configs` and `probe`. Every tooling command runs through the `sheets-framework` bin against the spreadsheet in its package's `sheets.config.json`.

One file per subject. Open the file the task needs.

| When | File |
| --- | --- |
| The Apps Script host, `triggerOnEdit`, the rollup preset, the ESLint preset | [apps-script-host.md](./how-it-runs/apps-script-host.md) |
| The Node host, what it leaves unproven, its HTTPS transport | [node-host.md](./how-it-runs/node-host.md) |
| The `sheets-framework` bin and `sheets.config.json` | [bin.md](./how-it-runs/bin.md) |
| The dev project and its fixture | [dev-project.md](./how-it-runs/dev-project.md) |
| What `gen-configs` writes, `configs-diff` | [gen-configs.md](./how-it-runs/gen-configs.md) |
| Running a chore, its dry run and preview | [chore-dry-run.md](./how-it-runs/chore-dry-run.md) |
| An auth error from the Node host, re-minting the token | [auth.md](./how-it-runs/auth.md) |
| Seeing the raw Sheets JSON with `probe` | [probe.md](./how-it-runs/probe.md) |
| What the `gworkspace` MCP can see | [gworkspace-mcp.md](./how-it-runs/gworkspace-mcp.md) |
