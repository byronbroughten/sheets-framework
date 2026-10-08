# Two ways to address a column

Map fragment. Sibling headings live in this folder.


A column can be named **relatively** — a `<TN, CN>` pair, Table name plus column name — or **absolutely**, by its full name (`"occupancy_updateTermsSelect"`) as a single correlated key. Both families are current; neither is being migrated away.

| | Relative `<TN, CN>` | Absolute `ColumnFullName` |
| --- | --- | --- |
| Shape | two parameters, uncorrelated | one key |
| Resolved by | distributing over the Table name | indexed access on the flat column map |
| Used by | every Named-tier class, and the operators above them | endpoint dispatch, and any type that names a *set* of columns |
| Filterable subsets | `ColumnNameFiltered<TN, VN, IF>`, within one Table | `ColumnFullName<VN, IF>`, across all Tables |

**Use relative addressing when the Table is already fixed by the receiver** — `table.column(name)`, `row.cell(name)`. Restating the Table at every one of those call sites would be a regression in ergonomics for no gain.

**Use absolute addressing when you need to name a set of columns that cuts across Tables** — "every column whose value name is `checkbox` and which isn't a formula" is inexpressible as a `<TN, CN>` pair, because the constraint doesn't live within one Table. That set is what lets `CheckboxColumnOperator` be attachable only to a real checkbox column, and what makes the endpoint-name unions checkable.

Both parameters of `ColumnFullName<VN, IF>` default to "don't care," so the bare form still denotes every column. `IF` deliberately does **not** default to `false`: that would silently shrink the union endpoint dispatch is keyed on. From a full name, `SheetNameOf`/`ColumnNameOf`/`ValueNameOf`/`ValueOf` recover the parts as plain indexed lookups — which works only because the flattening helper injects each column's Table name and column name *into* its own entry rather than merely encoding them in the key. Reaching them through a side map instead needs an intersection that quietly degrades every result to `any`, so the assertions in `SpreadsheetSchema.test.ts` are identity-based, not assignability-based.

**There is no type-level bridge from the relative pair to the absolute key**, and that is deliberate rather than missing — see [Type-check cost](./type-check-cost.md) for the measurement that rules it out, and [`docs/design.md`](../design.md)'s "Record a deliberate absence as deliberate" for why it's written down instead of left to look like an oversight. The two families meet only at runtime, where `ColumnSchema.fullName` builds a full name from a sheet and column name it already holds.

