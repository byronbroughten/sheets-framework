# Framework test style

Disclosed from [`docs/code-style.md`](../code-style.md), "Tests". How tests run, and the fakes: [`docs/testing.md`](../testing.md). The general test reasoning is in `@byronbroughten/config`'s `docs/code-style/tests.md`.

## Tier imports

The rule in `docs/code-style.md` holds a test to the downward-only tier-import rule production code follows, since a test co-located with its subject sits in that subject's tier. A test that imported from a higher tier would tie a lower tier's behaviour to code built on top of it.

## Endpoints through EndpointRun

The rule in `docs/code-style.md` tests an endpoint through `EndpointRun`, never by calling its action. Going through the run is what buys the selector pruning, the setup flush and the interplay between a wipe and the appends that follow it, all of which a rebuild-from-scratch endpoint depends on; the cost is a larger fixture, which is the right trade. For the same reason no test reaches for a private helper, a comparator or an intermediate list of lines. The seam and its fixtures: [`docs/testing.md`](../testing.md#testing-an-endpoint-through-endpointrun).

## Grid state over requests

The settled rule in `docs/code-style.md` asks a test of a write to assert the cells, rows, tabs and Tables it leaves, not the `batchUpdate` requests that carry them. It is the framework's case of the general rule on outcomes over calls. The same grid can come from different requests: one `updateCells` or three, an append or a reused blank row. A request assertion pins the route, so a refactor that merges or reorders requests fails tests while the sheet comes out the same.

Three kinds of test keep request assertions, because the requests are their subject. The first is `GoogleSheetsAPI.test.ts`, which tests the mapping from local operations to Google requests. The second is a round-trip count, which the test's name states. The third is the dry run's promise that a write puts nothing on the wire. The fake exposes only `batchUpdateCount()`, so the last two read the count. Lint flags a `batchUpdate` any other test supplies in an object literal; a genuine exception carries an inline disable naming its reason.

The fake replays every request kind onto its fixture and hands the test a `grid` to read ([`docs/testing.md`](../testing.md#which-requests-the-fake-replays)). When a test needs a request kind or field the fake does not replay, extend the replay rather than adding a request assertion.

## Real schema data over invented literals

A draft rule, proposed like the general ones. Use real, already-committed schema data over invented literals wherever the code under test resolves identifiers through the actual schema (column IDs via `columnConfigs.sheetConfig.x.columnId`, real sheet gids). This stops a test from passing against a shape that doesn't exist in production.

## Exemplar columns for type-level tests

The settled rule in `docs/code-style.md` asks a type-level test for an exemplar column whose value name can't churn under `gen:configs`. A config sheet's column qualifies, and so does a column of a dev fixture Table, such as `valueTypes` or `item` ([testing](../testing.md#the-dev-fixtures-and-their-exemplar-columns)).
