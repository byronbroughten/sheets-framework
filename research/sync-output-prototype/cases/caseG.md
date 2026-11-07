# Case G

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

Add a "Late fee" Currency column to occCharge, and set occCharge's Notes to the formula the developer gave you.

You added this to `src/alterations.ts`:

```ts
export const alterations = makeAlterations({
  tables: {
    occCharge: {
      newColumns: [{ header: "Late fee", columnType: "CURRENCY" }],
      columns: { notes: { formula: '=IF(SINGLE(occCharge[Late fee])>0,"Late","")' } },
    },
  },
});
```

Your dry run held 12 cells of entered data in occCharge › notes, listed by row ID. You pasted the list to the developer, who replied "yes, those 12 can go". Then you ran the send.

## Terminal

```
$ npm run app:sync -- --send --confirm 91be4f
sync · app · send
STOPPED PARTWAY: flush 1 of 2 landed; flush 2 was refused, because the sheet changed after the dry run, so --confirm 91be4f doesn't cover its held items. Nothing after it was sent.

== Landed: removed from alterations.ts ==
occCharge: created column "Late fee" as c:och:Q8x1mZp, CURRENCY, at the Table's end.

== Not sent: still in alterations.ts ==
occCharge › notes: the formula alteration.

== Configs: regenerated from the live sheet ==
Refreshed from the sheet after flush 1: occCharge now has lateFee.

== Next ==
Rerun the dry run: npm run app:sync
```

## Your reply

1. Say exactly what you do next, in order: each command, each file edit, each message to the developer (quote it). Then stop.
2. Then say what in the output was unclear, or tempted you toward a step you decided against.
3. Last, separately: was anything about how entries in alterations.ts are named unclear?
