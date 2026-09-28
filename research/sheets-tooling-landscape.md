# Which public tools already cover agent-driven Sheets API work?

Research for [#25](https://github.com/byronbroughten/sheets-framework/issues/25) (map [#24](https://github.com/byronbroughten/sheets-framework/issues/24)). Gathered 2026-09-27 from docs, source code (shallow clones of each repo's default branch), the GitHub API, npm and pypistats.

## Summary

Reaching the whole `batchUpdate` surface is common now, and so is Tables support. No tool previews a write against the sheet's state. Where a "dry run" exists, it prints the HTTP request it would send. The Sheets API has no `validateOnly` field, so nothing upstream can check a batch without applying it.

## The API baseline

- `spreadsheets.batchUpdate` takes `requests[]`, `includeSpreadsheetInResponse`, `responseRanges[]`, `responseIncludeGridData` and `commentsViewMode`. There is no validate-only or dry-run field. "If any request is not valid, no requests will be applied" ([reference](https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets/batchUpdate)).
- Tables are ordinary batchUpdate requests (`addTable`, `updateTable`, `deleteTable`), and `appendCells` takes a `tableId`. `@googleapis/sheets@18.0.1` `build/v4.d.ts` types them (`Schema$Request.addTable`, `deleteTable`, `updateTable`, and `tableId?` on several schemas). Any passthrough tool can therefore reach tables if the agent already knows the `tableId`.

## Tools

| Tool | Language | batchUpdate reach | Preview before write | Tables | Agent ergonomics | Traction (2026-09-27) |
| --- | --- | --- | --- | --- | --- | --- |
| **Google Sheets MCP server** (first-party, `sheetsmcp.googleapis.com/mcp/v1`) | remote service | Full. `update_spreadsheet` accepts "valid spreadsheets.batchUpdate Request object[s]" and lists 60+ types, including `addTable`/`updateTable`/`deleteTable` | None. Annotated as destructive and non-idempotent | Through raw requests. No table discovery or table-append helper | 6 tools: `get_values`, `get_spreadsheet`, `update_values`, `update_formulas`, `update_spreadsheet`, `insert_dimension`. Remote over OAuth, no local scripting | Workspace Developer Preview Program. Docs updated 2026-09-18 |
| **Google Workspace CLI** `gws` ([googleworkspace/cli](https://github.com/googleworkspace/cli)) | Rust binary, also on npm | Full. The command surface is generated at runtime from the Discovery Service, so `gws sheets spreadsheets batchUpdate --json …` works | `--dry-run` only echoes `{url, method, query_params, body}` without calling the API (`executor.rs:416-431`). It shows no state and no diff | Only by writing raw `addTable`/`appendCells` JSON. `+append` uses `values.append` (`helpers/sheets.rs:129`), which is not table-aware | Good: JSON in and out, `gws schema sheets.<method>`, 100+ bundled SKILL.md agent skills | 31.2k stars; 58k npm downloads/week. Last release v0.22.5 on 2026-03-31, with no main-branch commit since. "Not an officially supported Google product", pre-1.0 |
| **taylorwilsdon/google_workspace_mcp** (`workspace-mcp`) | Python | Curated only. About 20 Sheets tools, each building its own batchUpdate. No generic passthrough | None. It has `--read-only` and per-service permissions | Yes: `list_sheet_tables` and `append_table_rows`. The latter resolves the `tableId`, then sends `appendCells` with `tableId` (`gsheets/sheets_tools.py:1684-1775`) | MCP plus a `workspace-cli` that calls tools against a running server. Tool tiers keep context small | 3.2k stars; 63k PyPI downloads/week. v1.29.0 released 2026-09-27, very active |
| **xing5/mcp-google-sheets** | Python | Full. The `batch_update(spreadsheet_id, requests)` tool is a raw passthrough (`server.py:1435`), plus value helpers | None | Through raw requests only | MCP only, about 20 tools | 1.0k stars; 27k PyPI downloads/week. Last release v0.6.3 on 2026-05-14 |
| **a-bonus/google-docs-mcp** | TypeScript | Curated. Many Sheets tools, no generic passthrough | None | Yes: `createTable`, `getTable`, `listTables`, `appendTableRows`, `updateTableRange`, `deleteTable` (`src/tools/sheets/`) | MCP only, aimed at Claude Desktop | 664 stars. v1.11.3 released 2026-07-27, active |
| **gemini-cli-extensions/workspace** | TypeScript | Values only: `sheets.getMetadata`, `getRange`, `getText`, `write` | None | No | A Gemini CLI extension | 642 stars. Preview releases, the latest 2026-09-21 |
| **gspread** | Python | Full. `Spreadsheet.batch_update(body)` is a "lower-level method that directly calls" batchUpdate with an untyped dict (`spreadsheet.py:92`) | None | No table helpers | A human-oriented library with no agent affordances | 7.5k stars; 7.3M PyPI downloads/week. Last release v6.2.1 on 2025-05-14 |
| **google-spreadsheet** (npm, theoephraim) | TypeScript | Partial. The internal `_makeBatchUpdateRequest(requests: any[])` is used by row, cell and sheet helpers | None | No (no `tableId`/`addTable` in `src/lib`) | A row/cell object model for humans | 2.5k stars; 450k npm downloads/week. v5.3.0 released 2026-06-03 |
| **googleapis / @googleapis/sheets** | TypeScript | Full and fully typed: `sheets_v4.Schema$Request`, Tables types included | None | Types only. No discovery or helpers | The raw client. The agent must handle auth, IDs and ranges | googleapis 11.6M/week (v182.0.0); @googleapis/sheets 1.58M/week (v18.0.1); repo 12.3k stars |

## The gap the harness would fill

1. **A real preview.** Nothing resolves a batch against a snapshot to show what it will change, such as the rows, table ranges or sheets it touches, before the one-shot, non-validatable send. `gws --dry-run` prints the request JSON, which is the closest any tool comes.
2. **Typed scripting in the agent's loop.** No tool pairs a CLI that runs an inline TS script or JSON op batch with the same typed library that script imports. The typed options are libraries without agent ergonomics (googleapis). The agent-friendly ones take untyped JSON (gws, the MCP servers).
3. **Table awareness on top of the full passthrough.** Tables support exists in curated MCP servers (taylorwilsdon, a-bonus). Full passthrough exists in gws, xing5 and Google's server. None has both: discovering table IDs, table-anchored helpers that compile down to requests, and any raw request besides.
4. **Snapshot as a safety net and per-target send policy.** No tool takes a pre-send snapshot or gates a send per spreadsheet. The nearest are coarse `--read-only` modes.

So the gap is real but narrow, and it lies in preview, snapshots, typed inline scripting and table-aware helpers. Reaching batchUpdate is solved. If the harness is ever published, gws is the closest overlap and its obvious comparison. Its momentum has stalled since March 2026, but it has far more stars.
