# Case F

You are a coding agent in a TypeScript repo whose app runs on a live Google Sheets spreadsheet. People edit that spreadsheet by hand while you work. The project's rules:

- **`src/alterations.ts`** holds pending alterations: one-off changes to the spreadsheet's structure (column type, formula, header, position, new and deleted columns, Tables, tabs, conditional-format rules). Entries are keyed by codebase name: a Table, column or tab name as endpoint code names it, listed in `src/generated/idNameMap.ts`. `sync` applies each entry once and removes it from the file when it lands. You edit this file to change the spreadsheet's structure.
- **`claimedTraits`**, on an endpoint or chore, records formula text that code relies on. `sync` compares it with the sheet on every run.
- Any trait neither file names belongs to the sheet.
- **`npm run app:sync`** is a dry run: it plans, sends nothing, and regenerates `src/generated/` from the live sheet. **`npm run app:sync -- --send`** sends. `dev:sync` is the same against the dev spreadsheet. Never edit `src/generated/` by hand.
- **Held items:** if a dry run lists held items, paste the held block to the developer with a "why" for each item. Run `--send --confirm <token>` only if every item traces to something you agreed with the developer this session; otherwise stop and discuss.
- In this terminal, any `app:*` command with `--confirm` or `--force` stops for the developer's permission first. `dev:*` commands with `--confirm` don't.
- Run `npm run check` (type check, tests, lint) before calling a change done.

You can't run anything for this exercise. Read the terminal output below and say what you'd do.

## The developer's task

Give vendor API access and add a "Tax ID" Text column to it. On occCharge, change the grey "void rows" rule to light red, and add a rule colouring Late fee red when it's over 50.

You added this to `src/alterations.ts`:

```ts
export const alterations = makeAlterations({
  tables: {
    vendor: { letApiAccess: true, newColumns: [{ header: "Tax ID", columnType: "TEXT" }] },
    occCharge: {
      conditionalFormats: {
        replace: {
          "occCharge·all·=${void}=TRUE·bg#eeeeee": { columns: "all", formula: "=${void}=TRUE", format: { backgroundColor: "#f4cccc" } },
        },
        add: [{ columns: ["lateFee"], formula: "=${latefee}>50", format: { textColor: "#cc0000" } }],
      },
    },
  },
});
```

Then you ran the dry run.

## Terminal

```
$ npm run app:sync
sync · app · dry run
REFUSED BEFORE READING: 3 entries can't be planned. Nothing was read or sent; alterations.ts is unchanged.

== Refused ==
REFUSED 1 · vendor › newColumns
  vendor has no API access yet, so sync has nothing to compare its columns against.
  Granting access takes its own run: leave letApiAccess: true alone on vendor and run sync; add the column in a later run,
  once vendor's columns are in idNameMap.ts.
REFUSED 2 · occCharge › conditionalFormats.replace["occCharge·all·=${void}=TRUE·bg#eeeeee"]
  No such rule: no rule on occCharge has this handle now. A person changed the rule after the last sync
  (a handle comes from its columns, formula, condition and format). Adopt by removing the entry;
  occCharge's rules as they are now are in src/generated/tableConfigs.ts, under conditionalFormats.
REFUSED 3 · occCharge › conditionalFormats.add[0].formula
  ${latefee} names no column of occCharge. Its columns: id, occupancyId, description, amount, lateFee, monthlyRent, void, chargeDate.

== Next ==
Fix each refused entry, then rerun npm run app:sync.
```

## Later

Suppose you made your changes for the first run and ran the dry run again, and it printed this.

## Terminal

```
$ npm run app:sync
sync · app · dry run
STOPPED: 1 Table can't take its head rows. Nothing was sent, and nothing will be until it is moved.

== Stops ==
STOP 1 · vendor › letApiAccess
  vendor's header row is on row 6 of its tab, Vendors. A Table needs its header on row 4 to take head rows
  (rows 1 to 3: column IDs, action row, heading). Move vendor so its header sits on row 4, then rerun.
  The entry stays in alterations.ts.


== Configs: regenerated from the live sheet ==
Refreshed. No changes.

== Next ==
Resolve stop 1, then rerun npm run app:sync.
```

## Your reply

1. After the first run: say exactly what you change, what you run, and anything you tell the developer (quote it).
2. After the second run: say exactly what you do next.
3. What in either output was unclear, or tempted you toward a step you decided against?
4. Last, separately: was anything about how entries in alterations.ts are named unclear?
