# Rules for `src/01_SpreadsheetSchema/`

- **Regenerate the config files, `appConfigs.ts` and `idNameMap.ts` in a package's `generatedDir` together with `sync`, never a subset**, and never edit their data by hand: fix the live tab, or write an `alterations.ts` entry.
- **A sheet-shape bug is fixed on the sheet, then regenerated.** Read the regenerated entry before building on it.
- **Fix a Table's spelling before code references it**: its name becomes the Table's key in the generated files and every string literal naming it.
- **`makeValueConfigs` keeps its `const` type parameter.** Without it every dropdown column's value type widens to `string`, silently.
- **Leave `customDefaultValue` at its default**: regeneration always emits it there, and no config column declares it.
- **Don't change `tableLayout.idDelimiter`**: every column ID already minted would stop parsing.
- **Refresh with a bare `sync` before editing `alterations.ts`**, so its entries name what the sheet holds now.
- **After a `sync` whose report lists a name that now points at another ID, reread `idNameMap.ts` before writing another `alterations.ts` entry**: a name read before it may name another Table, tab or column now.
- **Adopt a stop's live value freely, unless it conflicts with the task**; then ask.
- **`--force` and `--confirm` each need the developer's yes: ask by pasting the held or stopped block, with a why for each item.** The tier gate's ask on them: [byro-repo#564](https://github.com/byronbroughten/byro-repo/issues/564).
- Mechanics: [`docs/generated-data.md`](../../docs/generated-data.md).
