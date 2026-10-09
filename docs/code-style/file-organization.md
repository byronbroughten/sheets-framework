# Imports and file organization: examples

Disclosed from [`docs/code-style.md`](../code-style.md), "Imports & file organization". The rules are there, one line each; this file holds the examples. The general reasoning is in `@byronbroughten/config`'s `docs/code-style/file-organization.md`.

## The framework's public entries

`src/index.ts` is the Apps Script entry point, not a re-export barrel. The framework's two public entries are the only barrels: `src/framework.ts` (`@byronbroughten/sheets-framework`) and `src/frameworkTesting.ts` (`@byronbroughten/sheets-framework/testing`). An export joins them only when app code uses it. Every other file is imported directly by its path.
