# Candidate principles

Design reasoning for docs/design.md's "Not yet promoted" list. Candidates with one citation. Leave them here until a second decision makes the same argument; delete them if the first one gets reversed.

## A human-facing signal need not be machine-readable

A run's outcome is a cell background colour, which the read path never fetches — so no code can ever read a run's outcome back. That was acceptable because nothing did, and because the audience is a person looking at a sheet. A **conditional format outranks** that painted background, so a rule covering a feedback column can mask the yellow "started, never reported back" state while the colour is still written and still correct. That overlap is accepted rather than refused: the developer controls which rules they add, and turning run states into rules was rejected because a run state is a fact about one run at one moment, not a condition on a cell's value — the yellow case exists precisely because nothing wrote a conclusion.

_Cited by:_ #4, `ac7a795`, #45.

## A mechanism that needs an identity list waits for somewhere to keep it

The config-sheet floor's edit locks would keep other editors out only by naming editors explicitly, and there is no Admin Config to hold those addresses — so the floor ships warnings-only rather than half-fenced. An Admin Config also waits on why a lock naming the owner blocked the owner: an address list in the wrong form would lock the owner out and look exactly like Google ignoring the owner.

_Cited by:_ #49, #51.

## A tier is named for what it adds, and the layout sits below its first reader

The ladder is Source → Schema → Raw → Identified → Named: Source knows nothing about this spreadsheet, Schema is everything that reads the generated configs, Raw adds positional addressing on top of the layout and sheet list, Identified adds addressing by generated identity (sheet GID + column ID), and Named adds names. Identified was first called Indexed, which contradicted Raw being the positional tier, so "index" now means position everywhere. Raw reads the layout, so Schema sits below it; injecting the layout into Raw as props was rejected, since Raw has no use without one. Raw's column-blindness is a lint rule, not a convention.

_Cited by:_ #71, #72.

## Prefer the cheaper thing lazily over the complete thing eagerly

**Prefer the cheaper thing lazily over the complete thing eagerly**, when the complete version's cost scales with a union you don't control.

_Cited by:_ the lazy mapped-filter measurement in `docs/architecture/type-check-cost.md`.

## A fetched view plus a working view

Today's working view is the live sheet plus queued writes, indexes at pre-flush positions. Splitting those would let the fetched table stay put until flush while the working table moves immediately — the TODO on `RowRaw.delete`. Out of scope for the queue split.

_Cited by:_ #55, #58.

## Google Sheets stays behind one folder before a second platform exists

An Excel (Office Scripts) implementation should replace `src/00_Source/GoogleSheets/`, not chase Google types through the tiers. A platform capability reaches neutral code as a neutral type (`SheetChange`) or as a value the Apps Script host acts on (the `FloorNotice` `AppsScriptApi.handleSheetChange` shows), and a lint rule, not a convention, holds that.

_Cited by:_ #85, #104.
