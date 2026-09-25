# Framework test style

Disclosed from [`docs/style.md`](../style.md), "Tests". How tests run, and the fakes: [`docs/testing.md`](../testing.md). The general test reasoning is in `@byronbroughten/config`'s `docs/style/tests.md`.

## Real schema data over invented literals

A draft rule, proposed like the general ones. Use real, already-committed schema data over invented literals wherever the code under test resolves identifiers through the actual schema (column IDs via `columnConfigs.sheetConfig.x.columnId`, real sheet gids). This stops a test from passing against a shape that doesn't exist in production.

## Exemplar columns for type-level tests

The settled rule in `docs/style.md` asks a type-level test for an exemplar column whose value name can't churn under `gen:configs`. A config sheet's column qualifies, and so does a column of the live `test` sheet.
