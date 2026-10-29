# Imports and file organization: examples

Disclosed from [`docs/code-style.md`](../code-style.md), "Imports & file organization". The rules are there, one line each; this file holds the examples. The general reasoning is in `@byronbroughten/config`'s `docs/code-style/file-organization.md`.

## The framework's public entries

`src/framework.ts` is published as `@byronbroughten/sheets-framework` and `src/frameworkTesting.ts` as `@byronbroughten/sheets-framework/testing`. An export joins them only when app code uses it; every other file is imported directly by its path.
