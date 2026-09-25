# Tier words: Source, Platform, Raw, Identified, Named

The elaboration behind [docs/vocabulary.md](../vocabulary.md), whose one-line entries are the rules. These are the **architecture** words; the operator-facing ones live in [CONTEXT.md](../../CONTEXT.md).

If you're renaming or relocating something and aren't sure which word fits, ask rather than guess: these distinctions were deliberately hashed out, and folder placement depends on them.

- **Source** — the tier below everything that knows this spreadsheet: the `RawSource` port, the platform module (`GoogleSheets/`: the Google adapter and the Apps Script wrappers), and the cell values that cross the port.
- **Platform** — the spreadsheet product: Google Sheets now, Excel via Office Scripts later. It is a different axis from **host** (Apps Script or Node), so "platform-neutral" is not "host-neutral": platform-neutral code imports nothing from the platform module (`src/00_Source/GoogleSheets/`) and names no `GoogleAppsScript.*` type, and a lint rule holds that boundary.
- **Raw** — positional. It addresses sheets, rows and cells by GID and index, knows the layout and sheet list through `SpreadsheetBaseSchema` and `SpreadsheetSchema`, and never resolves a column by name or `columnId`; a lint rule holds Raw files to that. Raw's job is positional spreadsheet access only — fetching whole rows/columns and sheet properties through `RawSource`.
- **Identified** — addresses a column by generated identity (sheet GID + column ID), knows each column's config (value name, default, blank coercion), and types values by the value-name union. A column's index isn't stored in config data; Identified resolves it live from the column's `columnId` against the live "columnId" row.
- **Named** — config-dependent, addresses things by sheet name / column name.
