# Imports and file organization: examples

Disclosed from [`docs/style.md`](../style.md), "Imports & file organization". The rules are there, one line each; this file holds the examples. The general reasoning is in `@byronbroughten/config`'s `docs/style/file-organization.md`.

## The framework's public entries

`src/index.ts` is the Apps Script entry point, not a re-export barrel. The framework's two public entries are the only barrels: `src/framework.ts` (`@byronbroughten/sheets-framework`) and `src/frameworkTesting.ts` (`@byronbroughten/sheets-framework/testing`). An export joins them only when app code uses it. Every other file is imported directly by its path.

## Utility bundle names

`Str.ts` exports `Str`, and likewise `Obj`, `Arr`, `Tim` and `Val`. `SerialDate` is the exception, named for its type because it's the one utility the framework exports to business code, which has its own copy of `Arr`.
