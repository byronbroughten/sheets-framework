# The `gworkspace` MCP tools

A `gworkspace` MCP server (`workspace-mcp` on PyPI, run with `uvx` at a pinned version, limited to Docs, Drive and Sheets) reads and writes Google Docs and Sheets directly, separately from `clasp`/Apps Script. It signs in as one dedicated Google account and sees only the files shared with that account, so granting access is a share, not a config change. Take `spreadsheet_id` from the package's `sheets.config.json`; `list_spreadsheets` finds any spreadsheet shared with the account too.

- **It returns cell values, and table and column names, but not a table's declared column types.** `tables[].columnProperties[].columnType` and a column's validation rule are invisible to `read_sheet_values` and `list_sheet_tables`. Reading those means calling the Sheets REST API with the `clasp` credential, and [`probe`](./probe.md) is how to do it: it is read-only, like a chore dry run.
