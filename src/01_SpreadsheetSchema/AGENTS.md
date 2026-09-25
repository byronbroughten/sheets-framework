# Rules for `src/01_SpreadsheetSchema/`

- **Regenerate the four files in `generated/` together with `npm run app:gen:configs`, never a subset**, and never edit their data by hand, floor entries included: fix the live tab or the seed.
- **A sheet-shape bug is fixed on the sheet, then regenerated.** Read the regenerated entry before building on it.
- **Fix a tab's spelling before code references it**: its title becomes the sheet's key in the generated files and every string literal naming it.
- **`makeValueConfigs` keeps its `const` type parameter.** Without it every dropdown column's value type widens to `string`, silently.
- **Don't repurpose `customDefaultValue`**: it is reserved for column defaults and not implemented; a read-time coercion belongs on the value schema.
- Mechanics: [`docs/generated-data.md`](../../docs/generated-data.md).
