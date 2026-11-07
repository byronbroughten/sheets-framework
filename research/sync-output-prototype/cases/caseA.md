# Case A

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

Make occPayment's Amount column Currency.

You added this to `src/alterations.ts`:

```ts
export const alterations = makeAlterations({
  tables: {
    occPayment: { columns: { amount: { columnType: "CURRENCY" } } },
  },
});
```

Then you ran the dry run.

## Terminal

```
$ npm run app:sync
sync · app · dry run
STOPPED: 1 conflict. Nothing was sent, and nothing will be until it is resolved.

== Stops ==
CONFLICT 1 · occPayment › amount › columnType · both changed
  alterations.ts (yours)        CURRENCY
  configs (last agreed)         AUTOMATIC
  live sheet                    NUMBER
  A person changed this on the sheet after the last sync, and your entry changes it too.
  Ways out:
    adopt  Remove the entry, or set it to NUMBER. Free, but your change to this trait is given up.
    force  Send CURRENCY over the sheet's NUMBER. Needs the developer's yes for this trait; the yes covers one run.
    ask    Change nothing and ask the developer which value they want.


== Plan: what --send would send once nothing stops it ==
 1. occPayment: set amount's type to CURRENCY (resends all 15 columns).
    Resets the style and colours of 4 dropdowns: payerCategory, nonOccupantPayerId, formOfPayment, paymentStatus (dropdownStyleReset: "allow").

== Configs: regenerated from the live sheet ==
Refreshed. Held back while conflict 1 is open: occPayment › amount › columnType stays AUTOMATIC in the configs, so the regen doesn't take up the sheet's NUMBER.

== Next ==
Resolve conflict 1, then rerun npm run app:sync.
```

## Your reply

1. Say exactly what you do next, in order: each command, each file edit, each message to the developer (quote it). Then stop.
2. Then say what in the output was unclear, or tempted you toward a step you decided against.
3. Last, separately: was anything about how entries in alterations.ts are named unclear?
