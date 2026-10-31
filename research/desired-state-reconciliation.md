# Research: how desired-state tools reconcile a sparse declaration against live state and a snapshot

For byronbroughten/sheets-framework#180 (part of #178). Retrieved 2026-10-31. Primary sources only; source-code links are pinned to the commit read:

- kubernetes/kubernetes `3c690c189e5d367715dbd056a5c2009a476c8d77`
- hashicorp/terraform `d40e1b2dfcd7ce883f2fddafa518a54a32a20a29`
- prisma/prisma-engines `d1b3f377ca835d73fd50954a5e22128b9ebdd6b8`
- sqldef/sqldef `e679c5c143080e18b7135b6889ec3dc60ee8c03f`

Anything not confirmed in a primary source is marked **unverified**.

## The design under test

The declarations are sparse and edited by agents. The generated configs are the last regen's snapshot of the live sheet, committed to git. The live sheet is the third input. Two rules:

- **Apply.** Declarations differ from the snapshot and live matches the snapshot.
- **Conflict.** Live differs from the snapshot and the declarations still hold the old value. Stop and name the property.

Properties the declarations omit belong to the sheet. Deletion is an explicit marker. An ID-prefix rename rewrites data.

---

## 1. Kubernetes client-side apply: `last-applied-configuration` (the closest analogue)

### The algorithm

`kubectl apply` (client-side) computes a patch from three documents: the configuration file (*modified*), the live object (*current*), and the `kubectl.kubernetes.io/last-applied-configuration` annotation on the live object (*original*). Every apply writes the configuration file into that annotation ([declarative-config, "Merge patch calculation"](https://kubernetes.io/docs/tasks/manage-kubernetes-objects/declarative-config/#merge-patch-calculation)). The documented steps:

1. **Fields to delete.** Fields in last-applied and missing from the configuration file. Fields set explicitly to `null` in the file are also cleared, "regardless of whether they appear in the `last-applied-configuration`".
2. **Fields to set.** Fields in the configuration file whose values don't match live.
3. Set the annotation to the configuration file.
4. Merge 1–3 into one patch request ([same section](https://kubernetes.io/docs/tasks/manage-kubernetes-objects/declarative-config/#merge-patch-calculation)).

For primitives the docs give a decision table ([primitive fields](https://kubernetes.io/docs/tasks/manage-kubernetes-objects/declarative-config/#merging-changes-to-primitive-fields)):

| In config file | In live | In last-applied | Action |
|---|---|---|---|
| Yes | Yes | – | Set live to config value |
| Yes | No | – | Set live to config value |
| No | – | Yes | Clear from live |
| No | – | No | Do nothing, keep live value |

Maps are merged key by key with the same table ([map fields](https://kubernetes.io/docs/tasks/manage-kubernetes-objects/declarative-config/#merging-changes-to-map-fields)). Lists use one of three strategies, chosen per field by the `patchStrategy` tag. A list with no strategy is replaced whole. A list of complex elements with `patchMergeKey` (for example `containers` keyed by `name`) is merged element by element ([list fields](https://kubernetes.io/docs/tasks/manage-kubernetes-objects/declarative-config/#merging-changes-for-fields-of-type-list)). In the merge-key example:

- An element missing from the file but present in last-applied is deleted.
- An element present only in live, and never in last-applied, "was retained".
- A live-only edit to a sub-field the file omits (`args`) is kept.

### The source, and the conflict mode nobody turns on

The engine is `strategicpatch.CreateThreeWayMergePatch(original, modified, current, schema, overwrite)` ([patch.go L2090–2177](https://github.com/kubernetes/kubernetes/blob/3c690c189e5d367715dbd056a5c2009a476c8d77/staging/src/k8s.io/apimachinery/pkg/util/strategicpatch/patch.go#L2090-L2177)). Its doc comment:

> "reconciles a modified configuration with an original configuration, while preserving any changes or deletions made to the original configuration in the interim, and not overridden by the current configuration … if overwrite is false and there are conflicts between the modified and current configurations. Conflicts are defined as keys changed differently from original to modified than from original to current."

It builds the patch in three steps:

- **`deltaMap`.** `diffMaps(current, modified)` with `IgnoreDeletions: true`.
- **`deletionsMap`.** `diffMaps(original, modified)` with `IgnoreChangesAndAdditions: true`.
- **Merge.** The two maps are merged into one patch.

When `overwrite` is false it also computes `changedMap = diffMaps(original, current)` and returns `ErrConflict` if `MergingMapsHaveConflicts(patchMap, changedMap)` ([L2157–2174](https://github.com/kubernetes/kubernetes/blob/3c690c189e5d367715dbd056a5c2009a476c8d77/staging/src/k8s.io/apimachinery/pkg/util/strategicpatch/patch.go#L2157-L2174)).

`kubectl apply` exposes this as `--overwrite`, **default `true`**: "Automatically resolve conflicts between the modified and live configuration by using values from the modified configuration" ([apply.go L188, L236](https://github.com/kubernetes/kubernetes/blob/3c690c189e5d367715dbd056a5c2009a476c8d77/staging/src/k8s.io/kubectl/pkg/cmd/apply/apply.go#L236)). The value is passed straight to `CreateThreeWayMergePatch` ([patcher.go L312](https://github.com/kubernetes/kubernetes/blob/3c690c189e5d367715dbd056a5c2009a476c8d77/staging/src/k8s.io/kubectl/pkg/cmd/apply/patcher.go#L312)). So by default client-side apply **silently overwrites** out-of-band changes to fields the file names. Only `--overwrite=false` gives our "stop on conflict" behaviour. (The separate `--force` path deletes and recreates the object on a 409 or Invalid error, [patcher.go L339–340](https://github.com/kubernetes/kubernetes/blob/3c690c189e5d367715dbd056a5c2009a476c8d77/staging/src/k8s.io/kubectl/pkg/cmd/apply/patcher.go#L339-L340). That is a different "conflict": a resourceVersion race.)

### The five questions

**Conflict surfacing.**
- With `--overwrite=false`, an `ErrConflict` carries both the patch and the original→current change as YAML.
- With the default, no conflict surfaces at all.
- The SSA docs name the cost: under client-side apply "outdated values which have been overwritten by other users are left in an applier's local config … an applier has no way of knowing whether their next apply will overwrite other users' changes" ([SSA, "Comparison with Client-Side Apply"](https://kubernetes.io/docs/reference/using-api/server-side-apply/#comparison-with-client-side-apply)).

**Override.** `--overwrite=true` (the default).

**Ownership hand-back.** To give a field from the file to an imperative writer, the docs require manual steps: remove it from the file **and** from the annotation ([change ownership](https://kubernetes.io/docs/tasks/manage-kubernetes-objects/declarative-config/#changing-the-owner-from-a-configuration-file-to-a-direct-imperative-writer)). Deleting it from the file alone would *clear* it, because it is still in last-applied.

**Omission vs deletion.** Omission means two different things depending on history:
- A field never declared is left alone (row 4).
- A field once declared and then dropped is **cleared** (row 3).

An explicit `null` always clears. Whole-object deletion is recommended as an explicit imperative `kubectl delete -f`, "more explicit about what is being deleted, and less likely to result in the user deleting something unintentionally". Deletion by absence (`--prune`) is opt-in and alpha or ApplySet-based ([How to delete objects](https://kubernetes.io/docs/tasks/manage-kubernetes-objects/declarative-config/#how-to-delete-objects)).

**Creation keying.** Objects are keyed by client-chosen kind/namespace/name, so nothing waits for a server ID. List elements are keyed by `patchMergeKey`.

**Defaults pitfall.** Server-defaulted fields are absent from last-applied, so apply never clears them. Changing `strategy.type` to `Recreate` leaves the defaulted `rollingUpdate` in place, and the apply fails. The docs recommend declaring such coupled fields explicitly, and clearing defaulted fields by setting them to `null` ([Default field values](https://kubernetes.io/docs/tasks/manage-kubernetes-objects/declarative-config/#default-field-values)).

**Renames.** None; a name change is a different object.

**Destructive gating.** None beyond `--dry-run` and `kubectl diff`.

## 2. Kubernetes server-side apply (SSA)

**Model.** The server records per-field ownership in `.metadata.managedFields`, one entry per manager, with an operation of `Apply` or `Update`. "Whenever a field's value does change, ownership moves from its current manager to the manager making the change", and managers setting the same value share ownership ([Field management](https://kubernetes.io/docs/reference/using-api/server-side-apply/#field-management)).

**Conflict surfacing.** "A *conflict* is a special status error that occurs when an `Apply` operation tries to change a field that another manager also claims to manage." Apply fails on conflict unless forced. A plain update (PUT) never fails on conflict and simply takes ownership ([Conflicts](https://kubernetes.io/docs/reference/using-api/server-side-apply/#conflicts); [apply and update](https://kubernetes.io/docs/reference/using-api/server-side-apply/#apply-and-update)). The conflict names the field and the other manager; "Managers identify distinct workflows … (especially useful on conflicts!)" ([Field managers](https://kubernetes.io/docs/reference/using-api/server-side-apply/#managers)).

**Three documented resolutions** ([Conflicts](https://kubernetes.io/docs/reference/using-api/server-side-apply/#conflicts)):
1. **Overwrite and become sole manager.** `force=true` / `--force-conflicts` "changes the value of the field, and removes the field from all other managers' entries".
2. **Give up the claim.** Remove the field from your manifest; the value is unchanged.
3. **Become a shared manager.** Change your manifest to match the live value; the value is unchanged and ownership is shared.

**Omission vs deletion.** "If you remove a field from a manifest and apply that manifest, Server-Side Apply checks if there are any other field managers that also own the field. If the field is not owned by any other field managers, it is either deleted from the live object or reset to its default value" ([Field management](https://kubernetes.io/docs/reference/using-api/server-side-apply/#field-management)). Omission therefore means "I no longer manage this". It deletes only when nobody else claims the field.

**Granularity markers.** `listType` (`atomic`/`set`/`map`) with `listMapKey`, and `mapType`/`structType` (`atomic`/`granular`), decide whether a list or map is owned whole or per entry ([Merge strategy](https://kubernetes.io/docs/reference/using-api/server-side-apply/#merge-strategy)).

**Migration from client-side apply.** The annotation is used to infer which fields client-side apply owns. "Any fields not managed by client-side apply raise conflicts", for example `replicas` changed by `kubectl scale` ([Upgrading](https://kubernetes.io/docs/reference/using-api/server-side-apply/#upgrading-from-client-side-apply-to-server-side-apply)).

**Creation, rename, destructive gating.** Creation is by name, as in section 1. There are no renames and no destructive classification.

## 3. Terraform

**Model.** State's "primary purpose … is to store bindings between objects in a remote system and resource instances declared in your configuration" ([State](https://developer.hashicorp.com/terraform/language/state)). By default `plan`:
- reads the current state of existing remote objects,
- compares the configuration to the prior state,
- and proposes actions to make remote objects match the configuration ([plan](https://developer.hashicorp.com/terraform/cli/commands/plan)).

Apply then "Updates the state file so that the configuration, real infrastructure, and state match" ([resource behavior](https://developer.hashicorp.com/terraform/language/resources/behavior)).

**Conflict surfacing: drift is reported, then overwritten.** Terraform is not three-way in our sense. After refresh, it prints "Note: Objects have changed outside of Terraform" and the drifted attributes. It then warns: "Unless you have made equivalent changes to your configuration, or ignored the relevant attributes using ignore_changes, the following plan may include actions to undo or respond to these changes" ([jsonformat/plan.go L342–346, L364](https://github.com/hashicorp/terraform/blob/d40e1b2dfcd7ce883f2fddafa518a54a32a20a29/internal/command/jsonformat/plan.go#L342-L364)). In the equivalence-test fixture, the plan that follows reverts the drifted `string` to the configured value ([drift_relevant_attributes/plan](https://github.com/hashicorp/terraform/blob/d40e1b2dfcd7ce883f2fddafa518a54a32a20a29/testing/equivalence-tests/outputs/drift_relevant_attributes/plan)). The human gate is reading the plan before `apply`, plus `-detailed-exitcode` (2 = changes present) for automation ([plan](https://developer.hashicorp.com/terraform/cli/commands/plan)).

**Override and accept.**
- **`-refresh-only`.** "creates a plan whose goal is only to update the Terraform state and any root module output values", which accepts drift into the snapshot without touching infrastructure. When config already matches the drift, Terraform suggests `terraform apply -refresh-only` ([plan](https://developer.hashicorp.com/terraform/cli/commands/plan); [plan.go L168–176](https://github.com/hashicorp/terraform/blob/d40e1b2dfcd7ce883f2fddafa518a54a32a20a29/internal/command/jsonformat/plan.go#L168-L176)).
- **`-refresh=false`.** "causes Terraform to ignore external changes, which could result in an incomplete or incorrect plan."

**`lifecycle`** ([lifecycle](https://developer.hashicorp.com/terraform/language/meta-arguments/lifecycle)):
- **`ignore_changes`.** The named attributes are considered when planning a `create` but "ignored when planning an `update`". They take index notation (`tags["Name"]`), and `all` means Terraform never proposes updates. In effect: "seed on create, then the attribute belongs to the remote side".
- **`prevent_destroy`.** "Terraform rejects plans that would destroy the infrastructure object associated with the resource." But "This rule doesn't prevent Terraform from destroying a resource if you remove its configuration": the guard lives in the declaration, so deleting the declaration deletes the guard.
- **`create_before_destroy`.** Opt-in, "because many remote object types have unique name requirements".
- **`replace_triggered_by`.** Forces replacement when referenced resources change.

**Creation keying.** Resources are keyed by configuration address. The remote ID is bound in state after apply. Existing objects are adopted with an `import` block (`to` = address, `id` = the provider's ID, which "must be known during the plan operation") ([import block](https://developer.hashicorp.com/terraform/language/block/import)). Import is recorded as import, not creation. Keeping the block is recommended "as an historical artifact" because it is idempotent. A mismatch between config and the imported object shows up as a planned update ([single-resource import](https://developer.hashicorp.com/terraform/language/import/single-resource)).

**Deletion.** By absence: apply "Destroys resources that exist in the state but no longer exist in the configuration" ([resource behavior](https://developer.hashicorp.com/terraform/language/resources/behavior)). The explicit `removed` block exists for the opposite intent. With `lifecycle { destroy = false }` it "remove[s] the resource from state without destroying the actual resource" ([removed block](https://developer.hashicorp.com/terraform/language/block/removed)).

**Renames.** Without help, a changed address is read "as an instruction to destroy the existing resource" and create a new one. A `moved { from, to }` block renames the binding in state before planning, so "Terraform does not destroy the resource" ([moved block](https://developer.hashicorp.com/terraform/language/block/moved); [refactoring](https://developer.hashicorp.com/terraform/language/modules/develop/refactoring)). HashiCorp recommends keeping all historical `moved` blocks and chaining them; "Removing a `moved` block is a breaking change" ([refactoring](https://developer.hashicorp.com/terraform/language/modules/develop/refactoring)). `moved` renames an *address*, not remote data. A rename that has to rewrite remote data is a provider-level in-place update or a `ForceNew` replacement.

**Destructive classification.**
- **Per attribute, by the provider.** `ForceNew` "indicates that any change in this field requires the resource to be destroyed and recreated", and the plan shows it as a replace ([SDKv2 schema behaviors](https://developer.hashicorp.com/terraform/plugin/sdkv2/schemas/schema-behaviors)).
- **Forced by the operator.** `-replace=ADDR` forces a replacement ([plan](https://developer.hashicorp.com/terraform/cli/commands/plan)).
- **Gate.** No built-in data-loss gate beyond plan review and `prevent_destroy`.
- **Perma-diffs.** `DiffSuppressFunc` handles "capitalization differences in string names, or logical equivalences in JSON values" ([schema behaviors](https://developer.hashicorp.com/terraform/plugin/sdkv2/schemas/schema-behaviors)).

## 4. Declarative database schema tools

### Atlas (declarative `schema apply`)

**Model.** "Atlas plans a migration to get from the current state to the desired state", prints the SQL and prompts for approval. In declarative mode "any object that exists in the target but not in the reference is dropped", so deletion is by absence ([declarative apply](https://atlasgo.io/declarative/apply)).

**Gating.**
- `--auto-approve` skips the prompt and is "not recommended for production".
- `--dry-run` prints without executing.
- `lint.review` sets when a human must approve: `ERROR` (only when lint reports errors), `WARNING`, or `ALWAYS`. With `ERROR`, apply "auto-approves changes only when linting passes and requires manual approval when destructive changes are detected" ([declarative apply](https://atlasgo.io/declarative/apply); [analyzers](https://atlasgo.io/lint/analyzers)).

**Destructive classification** ([analyzers](https://atlasgo.io/lint/analyzers)):
- **Destructive (data loss):** DS101 drop schema, DS102 drop table, DS103 drop non-virtual column.
- **Data-dependent (may fail on existing data):** MF101 unique index on duplicates, MF103 non-null column on a non-empty table.
- **Backward-incompatible:** BC101/BC102 rename table/column.

Lint "fails with exit code 1 in case destructive changes are detected" unless `destructive { error = false }`. A per-statement `-- atlas:nolint destructive` suppresses it, and the Pro `force` setting overrides nolint. The docs suggest pre-migration checks asserting a table or column is empty before the drop.

**Renames.**
- Atlas v0.22 added an interactive planner prompt, "Did you rename "users" column from "first_name" to "name""; "No" yields drop-and-create ([v0.22 blog](https://atlasgo.io/blog/2024/05/01/atlas-v-0-22)).
- On 2026-05-20 Atlas added declarative `renamed_from` (HCL) / `-- atlas:renamed_from <prev>` (SQL), because "The interactive rename prompt is a human-only contract: agents cannot answer it". A non-interactive agent's rename otherwise falls back to DROP + CREATE and loses data ([changelog: declarative schema renames](https://atlasgo.io/changelog/declarative-schema-renames)).
- Whether the annotation should be removed after the rename applies is not stated (**unverified**).

**Out-of-band conflicts.** Not a concept in declarative mode: current state is re-inspected each run and diffed two-way against desired. **Unverified** whether Atlas has any snapshot-based drift detection in declarative mode; none was found in the pages read.

### Prisma (`db push` vs `migrate`)

**`db push`.**
- Introspects the database and executes the needed changes with no migration files ([prototyping](https://www.prisma.io/docs/orm/prisma-migrate/workflows/prototyping-your-schema)).
- "If `db push` anticipates that the changes could result in data loss, it will: Throw an error; Require the `--accept-data-loss` option" ([prototyping](https://www.prisma.io/docs/orm/prisma-migrate/workflows/prototyping-your-schema); [CLI reference](https://www.prisma.io/docs/orm/reference/prisma-cli-reference)).
- `--force-reset` wipes and recreates.

**Data-dependent classification (source).** Prisma's destructive change checker evaluates each warning against **live row and non-null counts**. Examples ([warning_check.rs L98–192](https://github.com/prisma/prisma-engines/blob/d1b3f377ca835d73fd50954a5e22128b9ebdd6b8/schema-engine/connectors/sql-schema-connector/src/sql_destructive_change_checker/warning_check.rs#L98-L192)):
- **Table drop.** `NonEmptyTableDrop` returns no warning when the row count is 0 ("dropping the table is safe if it's empty"). Otherwise: "You are about to drop the `{table}` table, which is not empty ({rows_count} rows)."
- **Column drop.** `NonEmptyColumnDrop` is silent when the table is empty or the column holds only nulls. Otherwise: "…which still contains {value_count} non-null values."
- **Casts.** `RiskyCast` and `NotCastable` follow the same pattern for type changes.
- **Recreated columns.** `DropAndRecreateColumn` warns when data would be lost.

When counts are unavailable, a softer "if the table is not empty" message is used.

**Renames.** A renamed field by default produces a migration that will "`CREATE` a new column … `DROP` the existing column … and the data in that column". The fix is `migrate dev --create-only`, then hand-editing the SQL to `RENAME COLUMN` ([customizing migrations](https://www.prisma.io/docs/orm/v7/prisma-migrate/workflows/customizing-migrations)). There is no rename annotation in the schema (none found; **unverified** that none exists in Prisma 7/8).

**Drift.** `migrate dev` detects drift between migration history and the database ("Drift detected: Your database schema is not in sync with your migration history") and offers a reset ([prototyping](https://www.prisma.io/docs/orm/prisma-migrate/workflows/prototyping-your-schema)). Drift is resolved by resetting, not by naming a per-field conflict.

### Skeema (MySQL)

**Unsafe operations are refused by default.** `skeema push` "refuses to run any DDL" and `diff` aborts when any operation is unsafe. The unsafe list includes ([safety](https://www.skeema.io/docs/features/safety/); [options](https://www.skeema.io/docs/options/)):
- drop table,
- drop non-virtual column,
- type changes that could lose, truncate or reduce precision,
- charset changes,
- engine changes,
- routine drops.

**Override.** `--allow-unsafe` is meant to be given on the command line each time; enabling it permanently is "not recommended … especially in the production environment".

**Size-gated automatic allowance.** `safe-below-size` allows unsafe operations on tables strictly below a byte size. "Skeema always treats empty tables as size 0 bytes", so `safe-below-size=1` means "only empty tables" ([options](https://www.skeema.io/docs/options/)).

**Renames.** Not supported. They are read "as DROP-then-ADD operations" and so blocked as unsafe. The workaround is to run the RENAME outside Skeema, then `skeema pull` the live schema back into the repo ([requirements](https://www.skeema.io/docs/requirements/)). `pull` is Skeema's "accept the out-of-band change" path, its refresh-only.

### sqldef (mysqldef/psqldef/...)

**Model.** Two-way: export the current schema, edit it, run `--dry-run`, then `--apply`. Operations are idempotent. There is an offline mode comparing two SQL files ([README](https://github.com/sqldef/sqldef/blob/e679c5c143080e18b7135b6889ec3dc60ee8c03f/README.md)).

**Deletion.** By absence ("Remove the line to DROP COLUMN"). But "By default, DROP operations are skipped (safe mode)", and `--enable-drop` turns them on for tables, columns, indexes and the rest ([cmd-psqldef.md](https://github.com/sqldef/sqldef/blob/e679c5c143080e18b7135b6889ec3dc60ee8c03f/cmd-psqldef.md)). Skipped statements are emitted as `-- Skipped: ...` comments that are never executed ([database.go L285–289](https://github.com/sqldef/sqldef/blob/e679c5c143080e18b7135b6889ec3dc60ee8c03f/database/database.go#L285-L289)). The same doc's earlier `--apply` example prints DROPs without `--enable-drop`, which reads as stale; the source comment and the flag text agree on skip-by-default. A design note (`object-management.md`) proposes per-pattern `drop`/`drop_column` permissions because `enable_drop` "is all-or-nothing". It is a proposal, not shipped behaviour (**unverified** whether implemented).

**Renames.** A declarative `-- @renamed from=old_name` annotation on tables, columns, indexes and enum values compiles to `ALTER … RENAME` ([README](https://github.com/sqldef/sqldef/blob/e679c5c143080e18b7135b6889ec3dc60ee8c03f/README.md)).

## 5. Schema-as-code for no-code databases and Sheets

No first-party declarative or reconciling tool was found for Airtable, Notion, Baserow, NocoDB or Google Sheets. A web search turned up only hosting-oriented Terraform pages for NocoDB and Baserow. Community Terraform providers were not evaluated (**unverified**). What the vendors expose are imperative schema APIs. These details matter for our design:

**Notion.** Updating a data source's schema keys properties by "property names or IDs". "Properties set to null will be removed", which is an explicit null tombstone. Renaming is `name` inside the property config. Type changes change "how Notion reads its existing values" ([update a data source](https://developers.notion.com/reference/update-a-data-source)).

**Airtable.** Update field "Updates the name, description, and/or options of a field". The only documented option is a formula's expression ([update field](https://airtable.com/developers/web/api/update-field)). No type-change and no delete-field endpoint was found on the official pages (**unverified**; third-party mirrors say type cannot change after creation).

**Google Sheets.**
- `AddSheetRequest` lets the client choose `sheetId`: "if one is not set, an id will be randomly generated. (It is an error to specify the ID of a sheet that already exists.)" ([batchUpdate requests](https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets/request)).
- Developer metadata can attach to rows, columns, sheets or the spreadsheet. It "remains associated at those locations as the spreadsheet is edited", the API assigns `metadataId` if omitted, and lookup is by `DataFilter` ([developer metadata guide](https://developers.google.com/workspace/sheets/api/guides/metadata)). Exactly how column metadata behaves on insert, move and delete is not stated there (**unverified**).
- Sheets Tables carry `columnProperties` (`columnIndex`, `columnName`, `columnType`, `dataValidationRule`). A dropdown type requires a `ONE_OF_LIST` rule, and other types must not set one. `DeleteTableRequest` deletes "the entire table and the contents of the table" ([tables guide](https://developers.google.com/workspace/sheets/api/guides/tables)).
- Whether `UpdateTableRequest` can change a `columnType`, and what happens to existing values when it does, is not stated (**unverified**).

---

## 6. Cross-tool summary

| | Conflict with an out-of-band change | Override | Keying before a server ID | Deletion | Data-rewriting rename | Destructive gate |
|---|---|---|---|---|---|---|
| kubectl client-side apply | Only with `--overwrite=false`: `ErrConflict` when original→modified and original→current change a key differently | `--overwrite` (default true, silent) | Client-chosen name; `patchMergeKey` for list items | Absence of a once-applied field; explicit `null`; `delete -f` for objects | n/a | none |
| SSA | Apply fails, naming field and manager | `--force-conflicts`; drop the field; or match the live value to share | Client-chosen name | Omission drops your claim; deleted only if no other owner | n/a | none |
| Terraform | Drift note, then the plan reverts it | Edit config; `ignore_changes`; `-refresh-only` to accept | Config address; ID bound in state; `import` for existing objects | Absence destroys; `removed` block to forget | `moved` block (address only) | plan review; `prevent_destroy`; provider `ForceNew` |
| Atlas | none (two-way) | n/a | Object name | Absence drops | `renamed_from`; interactive prompt for humans only | lint DS/MF/BC + `lint.review`; `--auto-approve` |
| Prisma | Migration-history drift, then reset | `--accept-data-loss`, `--force-reset` | Name | Absence drops | Hand-edited SQL | data-loss warnings computed from live row and non-null counts |
| Skeema | none; `pull` accepts live | `--allow-unsafe` per run | Name | Absence drops | unsupported (drop+add, blocked) | unsafe list; `safe-below-size` (1 = empty only) |
| sqldef | none (two-way) | n/a | Name | Absence, but drops skipped unless `--enable-drop` | `@renamed from=` | drop skipped and emitted as a comment |
| Notion API | n/a (imperative) | n/a | Name or ID | Explicit `null` | `name` field | none |

---

## 7. Patterns worth borrowing, mapped onto the declarations design

1. **Our conflict rule is `CreateThreeWayMergePatch` with `overwrite=false`, scoped to declared keys.** Define conflict exactly as Kubernetes does: a declared key "changed differently from original to modified than from original to current". Snapshot→declarations and snapshot→live both changed and disagree: stop. Kubernetes' default of `overwrite=true` is the documented failure mode SSA was built to fix. Keep ours off with no config switch; make overriding a per-run, per-property act.

2. **Offer SSA's three resolutions, named, in the conflict message.** For each conflicting property print both values and the three ways out:
   - **force:** declare and re-run with a per-property force flag; this reverts the human.
   - **adopt:** edit the declaration to the live value, the "shared manager" path. This is the default an agent should take.
   - **release:** delete the property from the declarations, the "give up claim" path; it then belongs to the sheet.

   Skeema's guidance applies to force: a command-line flag each time, never a standing setting.

3. **A refresh-only path is needed and must be gated.** Terraform's `-refresh-only` and Skeema's `pull` accept out-of-band changes into the snapshot. `gen:configs` already does this for undeclared properties. For a *declared* property, a refresh must not overwrite the snapshot silently (see risk A).

4. **Rename and ID-prefix rename as a declarative `renamed_from` / `moved` marker, kept as history.** Atlas, sqldef and Terraform all moved to a declarative annotation, and Atlas says why: an interactive prompt is "a human-only contract: agents cannot answer it". Model the ID-prefix rename as `{ from: "OLD-", to: "NEW-" }` on the Table. Classify it as "rewrites data" (the `findReplace`), the one named exception in the gate. After it lands, keep it, or retire it once the snapshot shows the new prefix, following Terraform's "removing a moved block is a breaking change" warning. **Decide which.** A stale `from` with no live match must be a no-op, not an error, the same idempotence Terraform claims for `import`.

5. **Deletion as an explicit tombstone, never absence.** Notion's `null` and Kubernetes' explicit `null` are the precedent. Kubernetes client-side apply also shows the hazard of the alternative: there, removing a once-declared field *clears* it. Our rule ("omitted belongs to the sheet") equals SSA's "release ownership" and Terraform's `removed { destroy = false }`, which is safer. Deleting a column then needs `{ delete: true }`. Deleting the guard must never delete the column; Terraform's documented `prevent_destroy` hole is exactly that.

6. **Gate destructive changes on live data, not on operation type alone.** Prisma computes warnings from live row and non-null counts, so a drop of an empty column is silent. Skeema's `safe-below-size=1` means "empty tables only". Atlas suggests pre-checks that a column is empty. This fits #178's gate rule ("no non-formula data cell deleted or overwritten"): the preview should count the non-empty, non-formula cells a column delete, type change or `findReplace` touches. Zero lets an agent proceed; more than zero asks. Report the count in the message, as Prisma does.

7. **Classify into tiers like Atlas: destructive, data-dependent, backward-incompatible.**
   - **Destructive:** column delete with data; type change that can't round-trip.
   - **Data-dependent:** declaring `emptyValueAllowed: false` on a column that already has blanks, Atlas MF103's analogue.
   - **Backward-incompatible:** renames that break endpoint code; for us `tsc`, which retarget handles.

8. **Key new entities by a client-chosen key, then bind the server ID in the snapshot.** This is Terraform's address→state binding and Kubernetes' client-chosen names. For a new Table, Sheets lets the client choose `sheetId` (an error if taken). Columns can be keyed by header name, with developer metadata as a possible stable anchor (behaviour on column move **unverified**). Inserted columns need a declared position *relative to a named neighbour*, not an index: indices shift.

9. **Plan, apply, refetch, regenerate with a detailed exit code.** Terraform's `-detailed-exitcode` (0 none, 1 error, 2 changes) and `-out` saved plans give an agent a machine-checkable "the preview showed no data loss" that apply can be bound to. Apply exactly the previewed plan, or fail if live changed between preview and apply.

10. **Normalise before comparing.** Use a `DiffSuppressFunc` equivalent for values the sheet rewrites: formula casing and spacing, number-format aliases, validation-rule ordering.

## 8. Risks in our design these tools reveal

**A. The snapshot is last-*observed* live, not last-*applied* declarations.**

- **Where the two differ.** Kubernetes' "original" is the last applied *declaration* (the annotation); Terraform's state is last observed *live*, and Terraform does not detect conflicts. Our rule uses an observed snapshot as "original". It works only while, for every declared property, snapshot == the declaration as last applied.
- **How it fails.** Suppose any regen writes a live value V1 into the snapshot while the declarations still say V0: a conflict run that still regenerates, a regen from another branch, or a merge of two branches' generated configs. The next run sees declarations ≠ snapshot and live == snapshot, classifies it as an agent change, and **silently reverts the human**. That is exactly what #178 forbids.
- **Mitigation.** Either (a) a conflicting declared property blocks writing that property's snapshot value, or the whole regen, or (b) record the last applied declared value per property alongside the snapshot (a `last-applied` sidecar) and use it as "original" for declared keys. The SSA docs make the same point: client-side apply can't tell whether its next apply overwrites someone.

**B. A send that doesn't land as planned looks like a fresh agent change.** If Sheets normalises a declared formula or rule, live never equals the declaration, and every run re-applies it, a perma-diff. Worse, if a human then edits it, the snapshot→live change and the declaration differ from each other, giving a spurious conflict. Normalise (pattern 10), and after apply-refetch compare declared against live and report "applied but not as declared", rather than leaving it for the next run.

**C. Coupled properties and defaults.** The Kubernetes `strategy.type`/`rollingUpdate` failure has a Sheets analogue. Declaring a Sheets Table `columnType` of dropdown *requires* a `ONE_OF_LIST` validation rule, and other types must not have one. A sparse declaration that names only `columnType` will collide with an undeclared, sheet-owned validation rule. Treat such pairs as one atomic declared unit, SSA's `atomic` struct type.

**D. A removed guard.** If `letApiAccess` or a `delete: true` marker is simply dropped from the declarations, ownership passes back to the sheet. That is correct, but an agent can't then tell "released" from "accidentally deleted line". Terraform's `prevent_destroy` caveat shows the same: the guard lives in the file it guards. Consider reporting released properties in the run summary.

**E. Positional keys under insertion.** Kubernetes and SQL tools never key by position. If declarations or the snapshot key columns by index, an out-of-band column insert shifts every later column. A three-way compare would then report many spurious conflicts, or worse, matching ones. Key by stable identity (header or metadata), and make insert/delete explicit operations, as the ticket already intends.

**F. Branches and git.** The snapshot is committed. Terraform's state is a single shared, locked object, not merged across branches. Two branches each regenerating the snapshot and merging can produce a snapshot no run ever observed, which feeds risk A. Consider treating the snapshot as regenerate-only on master, and rejecting hand or merge edits to it.
