# Rules for `src/chores/`

- **A chore's dry run is always safe**: the Node host adapter suppresses its writes.
- **`--send` applies a chore to the spreadsheet the package's `sheets.config.json` names, so it needs a yes naming that chore, unless the root's gates give that spreadsheet a standing one.** A yes for one spreadsheet never covers another ([the dry run](../../docs/how-it-runs.md#the-chore-and-its-dry-run)).
- **Verify a dry run's preview before handing it over**: compare the rendered requests with what the chore was meant to do, and call out anything wrong or larger than intended.
- **A chore is a one-off job run from the terminal**, never imported by `src/index.ts`. Homes and conventions: [`docs/architecture/chores.md`](../../docs/architecture/chores.md).
