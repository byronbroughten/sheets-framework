# Comments: reasoning and examples

Disclosed from [`docs/code-style.md`](../code-style.md), "Comments". The general comment reasoning is in `@byronbroughten/config`'s `docs/code-style/comments.md`.

## The framework's navigation blocks

The shape of a navigation block is in the general reasoning file. Six framework files have one:

- `src/02_SpreadsheetRaw/SpreadsheetRaw.ts`
- `src/02_SpreadsheetRaw/TableRaw.ts`
- `src/04_SpreadsheetNamed/SheetNamed.ts`
- `src/05_Operators/ConfigCoordinator.ts`
- `src/05_Operators/ConfigSheetFloor/ConfigSheetFloorEditWarnings.ts`
- `src/06_API/EndpointRun.ts`

Copy one of them for the shape. The set is small on purpose: a new block is the exception, not the pattern, and every other comment stays one line.
