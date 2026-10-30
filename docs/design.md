# Design philosophy

Why this codebase is shaped the way it is, in the reasoning that generalizes past the decision that produced it; the agent tooling's arguments live with the repo's agent docs, not here.

One line per principle. A principle's argument and instances, where it has them, are one file away. Open a reasoning file only when you're changing the rule, or the rule's line doesn't decide your case.

## Reasoning files

| When | File |
| --- | --- |
| Two pieces of state could contradict each other | [`docs/design/unrepresentable-disagreement.md`](./design/unrepresentable-disagreement.md) |
| Choosing the level a fact is stored or exposed at | [`docs/design/concept-granularity.md`](./design/concept-granularity.md) |
| Choosing which frame an index or coordinate counts in | [`docs/design/concept-frame.md`](./design/concept-frame.md) |
| A state model with a real condition it can't express | [`docs/design/room-for-real-states.md`](./design/room-for-real-states.md) |
| An absence in an API response | [`docs/design/payload-is-not-grid.md`](./design/payload-is-not-grid.md) |
| Adding a call to Sheets or another costly operation | [`docs/design/one-chokepoint.md`](./design/one-chokepoint.md) |
| A claim about cost, types or live Sheets behaviour | [`docs/design/falsifiable-claims.md`](./design/falsifiable-claims.md) |
| Which case gets the unmarked name | [`docs/design/unmarked-common-case.md`](./design/unmarked-common-case.md) |
| A gap that looks like an oversight | [`docs/design/deliberate-absence.md`](./design/deliberate-absence.md) |
| Adding a column to a config sheet | [`docs/design/config-sheet-columns.md`](./design/config-sheet-columns.md) |
| Declaring a shape in code or reading it from the sheet | [`docs/design/structure-vs-identity.md`](./design/structure-vs-identity.md) |
## Principles

- **Make disagreement structurally impossible rather than validating against it.** When two pieces of state can contradict each other, choose a shape where the contradiction is unrepresentable, not a check that catches it after the fact.
- **Model state at the granularity the concept actually has.** Store and expose a fact at the level it's about, not the level the wire format or the storage medium delivers it at.
- **Work in the concept's own frame; convert to the platform's only at the boundary, and brand the platform's frame.** An index counts from the thing it's about, and the one conversion sits where requests leave.
- **Give the model room for the states that actually occur.** A real condition the model can't express gets misreported as one it can; before settling a model, ask which conditions have nowhere to go.
- **The payload is not the grid.** An absence in a response is not an absence in the world; where a wire format elides the empty case, repair it once at the boundary.
- **Funnel the expensive thing through one place.** Give the costly operation one chokepoint, so instrumenting or optimizing it covers everything and no new call site adds cost unseen.
- **Make the claim falsifiable before believing it.** Run the cheap check on a claim about cost, types or live Sheets behaviour, then record the number or the assertion beside the conclusion.
- **Give the common case the unmarked name.** The reflexive name is the one wanted most of the time, decided by counting call sites; the rare case costs exactly one word.
- **Record a deliberate absence as deliberate.** In an AI-assisted codebase the riskiest gap is one that looks like an oversight: an unexplained gap reads as a to-do and gets filled in; a documented one carries its reason and survives.
- **A live config-sheet column must identify a row or serve a human on that sheet.** Sampled generated traits do not earn a cell.
- **Structure is declared in code; identity is recorded from the sheet.** Headers, floor columns, head rows and ID shape are declared; GID, `tableId`, column ID and prefix are read.

## Adding a principle

Decisions are ADRs in [`docs/adr/`](./adr/); the issue, commit, probe and code-site citations left in the reasoning files are evidence their migration kept. The rule is byro-repo's [ADR 0020](https://github.com/byronbroughten/byro-repo/blob/master/docs/adr/0020-a-principle-may-rest-on-one-adr-and-a-reasoning-file-holds-only-what-no-adr-holds.md).

- **A principle may rest on one ADR from this repo, or from real-estate under the last rule, and earns its line only by deciding cases its ADRs don't name.**
- **A principle gets a reasoning file only for what no single ADR holds**: the argument across several ADRs, or a citable case beyond its one ADR. Otherwise its line links its ADR directly.
- **A new ADR that makes a principle's argument goes in that principle's reasoning file**, as a link plus one clause on how it makes the argument.
- **A new principle is proposed to the developer and added only once approved**, as one line here, plus a reasoning file in `docs/design/` when it has one.
- **A superseded ADR stops counting.** A principle it leaves with no ADR goes to the developer.
- **Each bullet under the deliberate-absence principle names the gap and its reason.**
- **A real-estate ADR may be cited, as an absolute GitHub URL because these docs are published**, and counts toward a principle only when its argument is framework-general.
