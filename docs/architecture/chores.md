# Chores

Map fragment. Sibling headings live in this folder.


A **chore** is a unit of work run from the terminal against the live spreadsheet — the permanent home for the jobs that used to be improvised one route at a time. Every one-off against the live sheet is a chore: not a scratch function in `src/index.ts`, not an ad-hoc Sheets client in `scripts/`, not a deploy-to-run. Raw Sheets JSON comes from `sheets-framework probe`. The chore runner and its dry run are documented in [`how-it-runs.md`](../how-it-runs.md); this is where one goes.

**One typed exported const per file, named after its file**, mirroring how an endpoint entry is written: a short `description` the runner prints, and an `action` that receives the spreadsheet and the run (`ChoreRun`: the spreadsheet ID the runner bound it to). The description is what makes a durable chore legible a year later. `Chore.ts` holds the type.

**Chores are found by folder and filename, not a registry.** Nothing dispatches on a chore's name except the person typing it, so a registry would add churn in a shared file for no checking benefit. That is the deliberate difference from endpoints, where the framework dispatches on the key and the registry earns its place.

**Homes, by how long the work lasts and whose spreadsheet it serves** (a package lists its own in `choreHomes`):

- **The app's `src/chores/oneOff/`** — transient chores, deleted in the commit that records their run. The folder is meant to empty; a chore left there goes stale against a sheet shape that no longer exists.
- **The framework's `src/chores/`** — its generic chores, kept and listed in every package: `addMissingColumnIds`, `fillMissingRowIds`, and `ensureConfigSheetFloor`. A package chore with one of their names stops the runner. A twice-a-year repair belongs here, or in a package's own durable home, rather than earning a column and a checkbox on a sheet.
- **The framework's `dev/chores/`** — the dev spreadsheet's own home. `buildDevFixtures` is the dev fixture's recipe (its data and builder sit in `devFixtures/`, which the runner does not list): it creates missing fixture tabs, and a rebuild is deleting the tab and rerunning. It refuses any spreadsheet but the dev one, and reads its layout from `dev/generated/`, which `dev:chore` installs.
- **The framework tiers** — anything generalizable. A job that wants a capability the framework does not model is the evidence that the capability is worth building.

**A bulk edit across many cells is `findReplace`, not a raw request and not a loop.** It is the framework's sanctioned sweep — rename a stored value across a column, correct a typo across a sheet, edit a formula's text in place — and [queued-writes](./queued-writes.md) has its fields, its ordering and the read it invalidates. Reach for it before `gatherRawRequest`, whose obligation to file an issue exists precisely because a capability like this one was missing.

**Chores are not tested**, and that is deliberate: a test for a one-off would be a second statement of the same thing, written by the same hand in the same hour, and its real check is the preview read before saying send. Durable chores are the arguable middle and are still skipped. A chore that no longer type-checks against the current configs surfaces under `npm run tsc`, which is the intended failure.

