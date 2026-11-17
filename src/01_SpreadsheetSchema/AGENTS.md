# Rules for `src/01_SpreadsheetSchema/`

- **Regenerate the config files, `appConfigs.ts` and `idNameMap.ts` in a package's `generatedDir` together with `sync`, never a subset**, and never edit their data by hand: fix the live tab, or write an `alterations.ts` entry.
- **A sheet-shape bug is fixed on the sheet, then regenerated.** Read the regenerated entry before building on it.
- **Fix a Table's spelling before code references it**: its name becomes the Table's key in the generated files and every string literal naming it.
- **`makeValueConfigs` keeps its `const` type parameter.** Without it every dropdown column's value type widens to `string`, silently.
- **Leave `customDefaultValue` at its default**: regeneration always emits it there, and no config column declares it.
- **Don't change `tableLayout.idDelimiter`**: every column ID already minted would stop parsing.
- **After a `sync` whose report lists a name that now points at another column, reread `idNameMap.ts` before writing another `alterations.ts` entry**: a name read before it may name the other column now.
- Mechanics: [`docs/generated-data.md`](../../docs/generated-data.md).
