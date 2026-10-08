# Architecture vocabulary

The layering words, used precisely and never loosely; the operator-facing words (endpoint, selector, run state) are [CONTEXT.md](../CONTEXT.md)'s.

One line per term. The elaboration is one file away. Open a reasoning file only when you're changing the rule, or the rule's line doesn't decide your case.

## Reasoning files

| When | File |
| --- | --- |
| Source, Platform, Raw, Identified, Named, `rowIndex` | [`docs/vocabulary/tiers.md`](./vocabulary/tiers.md) |
| Tier word, State, Base, Operator, Collaborator | [`docs/vocabulary/class-names.md`](./vocabulary/class-names.md) |
| `xConfigs` / `XConfig` / trait, the config-sheet floor | [`docs/vocabulary/config.md`](./vocabulary/config.md) |
| Profile, sampled facts, which class a member belongs to, Working, head rows | [`docs/vocabulary/profile.md`](./vocabulary/profile.md) |
| Schema, blank cells, `value` / `valueOrEmpty` / `valueNotEmpty`, checkbox, `SerialDate`, `emptyValueAllowed` | [`docs/vocabulary/values.md`](./vocabulary/values.md) |

## Tiers

- **Source** is the tier below everything that knows this spreadsheet: the `RawSource` port, the platform module (`GoogleSheets/`), and the cell values that cross the port.
- **Platform** is the spreadsheet product (Google Sheets); **host** is where the code runs (Apps Script or Node). Platform-neutral code imports nothing from `src/00_Source/GoogleSheets/` and names no `GoogleAppsScript.*` type.
- **Raw** is positional: it addresses a sheet by GID, a Table by its live `tableId`, and rows and columns by Table-relative index (`rowIndex` 0 is the first body row; head rows by role). It is blind to column config, resolving a column only from the Table's live head rows, and sheet coordinates appear only at gathering.
- **Identified** adds the committed column config, reached by column ID, and is typed by value name.
- **Named** adds names: a Table by Table name, a column by header, a sheet by title as a container. It is the end developer's API.
- **`rowIndex` is the body index**: 0 is the Table's first body row, and a row counted from the sheet's top is a `SheetRowIndex`.
- **Schema** (`01_SpreadsheetSchema`) is the declared layout plus every reader of the generated configs, and sits below Raw. The spreadsheet schema classes and the value schema are two unrelated families sharing the word.

## Class names

- **Every tier class name ends in its tier word**: state types, bases and Common classes alike. A tier root is `<Scope>Base<Tier>`, with no exceptions.
- **State members are named by scope**, plus the tier word where one instance holds more than one tier's state at that scope.
- **Base means a class root (`<Scope>Base<Tier>`) or an index origin (`Base0`/`Base1`), and nothing else.** What the framework supplies whatever the spreadsheet adds is **Framework**.
- **An Operator extends a Named base and adds methods suited to one data structure**, reaching its subject through a getter. Business Operators live in `src/businessEndpoints/BusinessOperators/`.
- **A Collaborator is a class a coordinator builds from its own props and reaches through a lazy getter**, named `<Subject><Role><Tier>`, in a subfolder named after its coordinator. Callers use the coordinator.

## Config

- **Config is the data describing the spreadsheet's own structure, generated from the live sheet**, in each package's `generatedDir`.
- **`xConfigs` is the whole map, `XConfig` is one entry's record, and a trait is one property of one record.** "Trait" never means a collection.
- **The config-sheet floor is guaranteed, not data to fix**: the four config sheets' own entries always come out the same on regeneration.

## Writes

- **The queue layer says *write*, paired with *fetch*** (`fetchQueue` / `writeQueue`); *update* names no operation, key or field, because every queued operation is an update.
- **A queue key equals its operation's `kind`, and a `kind` equals its type name without "Operation"**, in lower camel case; `WriteOperations` is a mapped type over the kinds, so the compiler holds it. The one exception is `raw`, the escape hatch's `RawWriteOperation`.
- **A per-kind gather method is `gather<Key>Operation`, the key it pushes to**, and the per-row and per-sheet queue entries are `RowWrites` and `TableWrites`.
- **A field about one row or one column is singular (`appendRow`, `deleteRow`); a queue entry's field holding many is plural (`fillCells`, `fillColumns`); a `WriteOperations` key stays its kind.** A queue's flag is a directive to the flusher, like a config literal's, so it reads as a verb.
- **`FillCellOperation` and `FillColumnOperation` stay separate** so the send order, column fills then cell writes, lives in the structure; a multi-column fill would be `FillRangeOperation`.

## Profile, working and head rows

- **A Table's or column's profile is its descriptive facts as the working state holds them** (column type, validation, head-row contents, and the facts sampled from the top data row), which config regeneration derives configs from. It is Raw-only (`table.profile`, `column.profile`). Identity and geometry (`tableId`, name, bounds) and everything you do to a Table or column are on the Table or column itself.
- **Sampled facts are the column-wide facts read from a column's top data row** (formula, number format, top value), held in `ColumnStateRaw.sampledFacts`.
- **Working means present in the working view, by fetch or queued write, with row indexes at their pre-flush positions** (`workingRows`, `cell.inWorking`). "Active" is retired.
- **A head row is a row of the Table's head, the header row or one fixed above it, reached by its head role; two roles may share a row.**

## Values

- **Nearly every cell is tri-state: `Value<VN>` includes `""`.** Code that branches on a value says what empty means. `checkbox` is exactly `boolean`; don't restore the blank to it.
- **`valueNotEmpty` throws on a blank, `valueOrEmpty` keeps `""`, and `value` is whichever the column's Empty value allowed box declares.** The plurals follow suit; only Named offers `value`/`valueArr`.
- **An in-app date is a `SerialDate`, never a JS `Date`.** Build one with `SerialDate.fromYmd` or `SpreadsheetNamed.today()`, and move it with `SerialDate.addDays`/`addMonths`, not `+ 1`.
- **`emptyValueAllowed` shapes the accessor, not the value type**, and is enforced at access, never at fetch or write.
