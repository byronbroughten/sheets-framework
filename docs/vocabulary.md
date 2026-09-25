# Architecture vocabulary

The layering words, used precisely and never loosely; the operator-facing words (endpoint, selector, run state) are [CONTEXT.md](../CONTEXT.md)'s.

One line per term. The elaboration is one file away. Open a reasoning file only when you're changing the rule, or the rule's line doesn't decide your case.

## Reasoning files

| When | File |
| --- | --- |
| Source, Platform, Raw, Identified, Named | [`docs/vocabulary/tiers.md`](./vocabulary/tiers.md) |
| Tier word, State, Base, Operator, Collaborator | [`docs/vocabulary/class-names.md`](./vocabulary/class-names.md) |
| `xConfigs` / `XConfig` / trait, `spreadsheetConfig`, the config-sheet floor | [`docs/vocabulary/config.md`](./vocabulary/config.md) |
| Meta vs primary, crossing views, which class a member belongs to, Active, Active facts | [`docs/vocabulary/meta-primary.md`](./vocabulary/meta-primary.md) |
| Schema, blank cells, `value` / `valueOrEmpty` / `valueNotEmpty`, checkbox, `SerialDate`, `emptyValueAllowed` | [`docs/vocabulary/values.md`](./vocabulary/values.md) |

## Tiers

- **Source** is the tier below everything that knows this spreadsheet: the `RawSource` port, the platform module (`GoogleSheets/`), and the cell values that cross the port.
- **Platform** is the spreadsheet product (Google Sheets); **host** is where the code runs (Apps Script or Node). Platform-neutral code imports nothing from `src/00_Source/GoogleSheets/` and names no `GoogleAppsScript.*` type.
- **Raw** is positional: it addresses sheets, rows and cells by GID and index, and never resolves a column by name or `columnId`. Lint holds Raw files to that.
- **Identified** addresses a column by generated identity (sheet GID + column ID), knows each column's config, and resolves the column's index live from the "columnId" row.
- **Named** is config-dependent and addresses things by sheet name and column name.
- **Schema** (`01_SpreadsheetSchema`) is everything that reads the generated configs, and sits below Raw. The spreadsheet schema classes and the value schema are two unrelated families sharing the word.

## Class names

- **Every tier class name ends in its tier word**: state types, bases and Common classes alike. A tier root is `<Scope>Base<Tier>`, with no exceptions.
- **State members are named by scope**, plus the tier word where one instance holds more than one tier's state at that scope.
- **Base means a class root (`<Scope>Base<Tier>`) or an index origin (`Base0`/`Base1`), and nothing else.** What the framework supplies whatever the spreadsheet adds is **Framework**.
- **An Operator extends a Named base and adds methods suited to one data structure**, reaching its subject through a getter. Business Operators live in `src/businessEndpoints/BusinessOperators/`.
- **A Collaborator is a class a coordinator builds from its own props and reaches through a lazy getter**, named `<Subject><Role><Tier>`, in a subfolder named after its coordinator. Callers use the coordinator.

## Config

- **Config is the data describing the spreadsheet's own structure, generated from the live sheet**, in each package's `generatedDir`.
- **`xConfigs` is the whole map, `XConfig` is one entry's record, and a trait is one property of one record.** "Trait" never means a collection.
- **`spreadsheetConfig` is a single record**, with no `spreadsheetConfigs` collection.
- **The config-sheet floor is guaranteed, not data to fix**: the four config sheets' own entries always come out the same on regeneration.

## Meta / primary

- **Meta / primary is an axis orthogonal to the tiers, not a fourth tier.** Primary deals in contents and takes the unmarked name; Meta is the structure's own shape and takes a `Meta` stem.
- **Crossing views takes exactly one word**: `meta` from primary, `primary` from Meta. A sheet's `column(name)` and a column's `sheet` stay in the view you're in.
- **A member belongs on the Meta class only if it acts on a uniform row, samples the top data row for a column-wide fact, or reads the table's own column properties.** Everything else is primary.
- **Active means present in the working view, with row indexes at their pre-flush positions**, not "exists on the sheet".
- **Active facts are the column-wide facts sampled from a column's top data row** (formula, number format type, top value), held in `ColumnStateRaw` for table columns only.

## Values

- **Nearly every cell is tri-state: `Value<VN>` includes `""`.** Code that branches on a value says what empty means. `checkbox` is exactly `boolean`; don't restore the blank to it.
- **`valueNotEmpty` throws on a blank, `valueOrEmpty` keeps `""`, and `value` is whichever the column's Empty value allowed box declares.** The plurals follow suit; only Named offers `value`/`valueArr`.
- **An in-app date is a `SerialDate`, never a JS `Date`.** Build one with `SerialDate.fromYmd` or `SpreadsheetNamed.today()`, and move it with `SerialDate.addDays`/`addMonths`, not `+ 1`.
- **`emptyValueAllowed` shapes the accessor, not the value type**, and is enforced at access, never at fetch or write.
