# Framework style

The framework's own code-shape rules, layered on the general style doc that ships with `@byronbroughten/config` (`docs/style.md` in that package; `config/docs/style.md` in the workspace). Read that first: everything there applies here too. Where things live is [`src/AGENTS.md`](../src/AGENTS.md) and [`vocabulary.md`](./vocabulary.md).

One line per rule. The reasoning and worked examples are one file away. Open a reasoning file only when you're changing the rule, or the rule's line doesn't decide your case.

## Reasoning files

| When | File |
| --- | --- |
| Writing a coordinator or an Operator, placing a member or a class | [`docs/style/class-shape.md`](./style/class-shape.md) |
| Naming a column, sheet, row or collaborator, a live read, or a method verb | [`docs/style/naming.md`](./style/naming.md) |
| A type-level assertion, a cast, a registry literal, a utility type | [`docs/style/type-modeling.md`](./style/type-modeling.md) |
| A guard, a blank-tolerant read, a `catch` | [`docs/style/error-handling.md`](./style/error-handling.md) |
| Writing or changing a test | [`docs/style/tests.md`](./style/tests.md) |
| A file-level navigation block | [`docs/style/comments.md`](./style/comments.md) |
| A barrel, or a utility bundle's file name | [`docs/style/file-organization.md`](./style/file-organization.md) |

## Class shape

- **Coordinating other stateful objects means a coordinator class extending the tier's Base class.** Callers build one with `init`; it builds its collaborators with `new` from props on `this`, as lazy getters. Endpoints are exempt until a body turns unwieldy.
- **An Operator extends its subject's `*BaseNamed` and reaches the subject through a getter**, never by extending the concrete class or taking one as a constructor argument.
- **What an Operator holds as props is its identity; a per-run value is an argument to the method that needs it.**
- **Split a coordinator into collaborators when its private helpers fall into groups that share nothing with each other**, not when it passes a method count. The coordinator keeps its public methods as one-line delegations, and the collaborators go in a subfolder named after it.
  - **Collaborators get no tests of their own** and reach the coordinator's shared surface through a getter. A base class stays in `ClassBases/`.
- **A composition of collaborator calls that answers one domain question belongs on the collaborator**, under its own name. A parameter that its only caller already holds as its own state means the query belongs on the instance.
- **A member that samples the top data row for a column-wide fact belongs on the Meta column.** `topCell`/`topRow` stay primary.

## Naming

- **Prefix a getter `active` when it reads live sheet state that has a same-named schema/config counterpart.** A helper that moves down onto the object it's about renames `_actualX` → `activeX`.
- **`column` abbreviates to `col` by default, and is spelled out beside an already-short suffix**, one form per scope.
- **A sheet takes the unmarked name and a row is marked with a spelled-out `Row` suffix.**
- **Google's API names stay at the wire; framework names follow the glossary.**
- **A method that deletes more than one row takes a `SHOUTING_SNAKE_CASE` name**, and keeps it once a guard makes the operation safe.
- **A collaborator is named `<Subject><Role><Tier>`, the role the agent noun of a verb on the list below.** A job with no verb on the list gets a plain descriptive noun, never "Handler" or "Manager".
- **Method names draw from one controlled verb vocabulary** — don't invent a new verb for a meaning already on this list:
  - `fetch` — actually hits the live Sheets API
  - `integrate` — merges a fetched snapshot into local state; no API call
  - `prep`/`gather` — queue state locally before a fetch (`prepFetchX` queues only; `gatherFetchX` queues _and_ fetches)
  - `update` — writes a local/queued change, not yet flushed
  - `append` — adds a new row
  - `ensure` — idempotent guard: make this true, no-op if it already is
  - `validate` — asserts an invariant, throws on failure
  - `init` — factory setup
  - `sync`/`flush` — coordinate multiple operators / send a batched write
  - `discard` — drop queued changes without sending them; the counterpart to `flush`
  - **The list governs framework methods.** A business operator's public method takes its verb from the app's `packages/real-estate/CONTEXT.md` instead.

## Comments

- **The framework's file-level navigation blocks are the six listed in the reasoning file**; a new one is the exception.

## Error handling & validation

- **`Val.assert(value, "label")` for "this shouldn't be missing" guards**, never a bare `!`; lint rejects `!`.
- **Read and validate in one step; reach for a marked read (`valueOrEmpty`, `valueNotEmpty`) only where the call site's requirement differs from its column's** Empty value allowed declaration.
- **`try`/`catch` has no established convention yet**; don't generalize from its one use.
- **`value`/`valueOrEmpty` throw on a row never fetched**, so a decision that branches on a cell queues its fetch in the same cycle, or says at the call site what an unfetched row means.

## Type modeling

- **Generic params use the framework's abbreviations** (`SN`, `VN`, `CN`, `UN`, `IF`, `TN`); its domain-free utilities are `utils/` and `appUtils/`.
- **Verify a type-level claim with `IsExactly` / `assertType` / `assertNotType` from `src/testSupport/typeAssertions.ts`, never an assignment.** Measure a mapped type over the config unions before adopting it.
- **`as` casts narrow data that's already runtime-safe; they never substitute for validation.** External values go through `Val.validate.*`/`Val.is.*`. The three accepted cast idioms are in the reasoning file.
- **A registry keyed by a finite name union takes a plain `: Type` annotation, not `makeStructuredConfig`**, which stays for the generated config files.
- **Custom generic utility types live in `utils/Obj.ts`**, PascalCase, one clear transform per name.

## Functional vs. imperative idioms

- **`for…in` appears only in `utils/`**, the framework's structural utilities.

## Tests

- **A test imports only from its own tier and below.**
- **Test an endpoint through `EndpointRun`, never by calling its action**, and assert the batch-update requests the run emits. No test reaches for a private helper.
- **A type-level test names an exemplar column whose value name can't churn under `gen:configs`.**

## Tooling

- **The framework's `scripts/` follow the general style and this file too**, all but the tier class-shape rules (coordinators, Operators, collaborators, Meta columns) and the Sheets verb list, and may use `Val`.

## Imports & file organization

- **The framework's two public entries, `src/framework.ts` and `src/frameworkTesting.ts`, are its only barrels**; `src/index.ts` is the Apps Script entry point, not a barrel.
- **Its utility bundles are `Str`, `Obj`, `Arr`, `Tim` and `Val`**; `SerialDate` is the exception, named for its type because it's the one utility the framework exports to business code, which has its own copy of `Arr`.
- **Tier subfolders**: `ClassBases/` for base and Common classes + their prop interfaces; `Types/`/`ClassTypes/` for supporting state/shape types consumed by that tier's classes.
