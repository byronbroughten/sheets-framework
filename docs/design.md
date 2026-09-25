# Design philosophy

Why this codebase is shaped the way it is, in the reasoning that generalizes past the decision that produced it; the agent tooling's arguments live with the repo's agent docs, not here.

One line per principle. Each principle's argument and instances are one file away. Open a reasoning file only when you're changing the rule, or the rule's line doesn't decide your case.

## Reasoning files

| When | File |
| --- | --- |
| Two pieces of state could contradict each other | [`docs/design/unrepresentable-disagreement.md`](./design/unrepresentable-disagreement.md) |
| Choosing the level a fact is stored or exposed at | [`docs/design/concept-granularity.md`](./design/concept-granularity.md) |
| A state model with a real condition it can't express | [`docs/design/room-for-real-states.md`](./design/room-for-real-states.md) |
| An absence in an API response | [`docs/design/payload-is-not-grid.md`](./design/payload-is-not-grid.md) |
| Adding a call to Sheets or another costly operation | [`docs/design/one-chokepoint.md`](./design/one-chokepoint.md) |
| A claim about cost, types or live Sheets behaviour | [`docs/design/falsifiable-claims.md`](./design/falsifiable-claims.md) |
| Which case gets the unmarked name | [`docs/design/unmarked-common-case.md`](./design/unmarked-common-case.md) |
| A gap that looks like an oversight | [`docs/design/deliberate-absence.md`](./design/deliberate-absence.md) |
| Adding a column to a config sheet | [`docs/design/config-sheet-columns.md`](./design/config-sheet-columns.md) |
| Declaring a shape in code or reading it from the sheet | [`docs/design/structure-vs-identity.md`](./design/structure-vs-identity.md) |
| A parked candidate under "Not yet promoted" | [`docs/design/candidates.md`](./design/candidates.md) |

## Principles

- **Make disagreement structurally impossible rather than validating against it.** When two pieces of state can contradict each other, choose a shape where the contradiction is unrepresentable, not a check that catches it after the fact.
- **Model state at the granularity the concept actually has.** Store and expose a fact at the level it's about, not the level the wire format or the storage medium delivers it at.
- **Give the model room for the states that actually occur.** A real condition the model can't express gets misreported as one it can; before settling a model, ask which conditions have nowhere to go.
- **The payload is not the grid.** An absence in a response is not an absence in the world; where a wire format elides the empty case, repair it once at the boundary.
- **Funnel the expensive thing through one place.** Give the costly operation one chokepoint, so instrumenting or optimizing it covers everything and no new call site adds cost unseen.
- **Make the claim falsifiable before believing it.** Run the cheap check on a claim about cost, types or live Sheets behaviour, then record the number or the assertion beside the conclusion.
- **Give the common case the unmarked name.** The reflexive name is the one wanted most of the time, decided by counting call sites; the rare case costs exactly one word.
- **Record a deliberate absence as deliberate.** An unexplained gap reads as a to-do and gets filled in; a documented one carries its reason and survives.
- **A live config-sheet column must identify a row or serve a human on that sheet.** Sampled generated traits do not earn a cell.
- **Structure is declared in code; identity is recorded from the sheet.** Headers, floor columns and ID shape are declared; GID, column ID and prefix are read.

## Not yet promoted

- **A human-facing signal need not be machine-readable.**
- **A mechanism that needs an identity list waits for somewhere to keep it.**
- **A tier is named for what it adds, and the layout sits below its first reader.**
- **Prefer the cheaper thing lazily over the complete thing eagerly**, when the complete version's cost scales with a union you don't control.
- **A fetched view plus a working view**, so the fetched table stays put until flush.
- **Google Sheets stays behind one folder before a second platform exists.**

## Adding a principle

- **Every principle cites the decisions that produced it**, by issue where one exists and by commit otherwise. A principle with no citation is a platitude and is cut.
- **A principle that can cite only one decision is parked under "Not yet promoted"** until a second decision makes the same argument, and deleted if the first one is reversed.
- **A new principle is one line here plus one reasoning file in `docs/design/`; a new instance is a sentence in that file.** A new candidate is one line here plus a heading in `candidates.md`.
- **There is no ADR tree.** Specs are published as GitHub issues, and a second filing system would only drift from them.
