# Case E

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

Add a "Late fee" Currency column to occCharge, right before Rent, and make unit's Move in date required.

Earlier this session you added this to `src/alterations.ts`, taking the names from `src/generated/idNameMap.ts`. People have been editing the sheet since.

```ts
export const alterations = makeAlterations({
  tables: {
    occCharge: { newColumns: [{ header: "Late fee", columnType: "CURRENCY", before: "rent" }] },
    unit: { columns: { moveIn: { emptyValueAllowed: false } } },
  },
});
```

Now you ran the dry run.

## Terminal

```
$ npm run app:sync
sync · app · dry run
READY: nothing held. This was a dry run: nothing was sent.

== Plan: what --send sends, in order ==
 1. occCharge: insert column "Late fee" (CURRENCY) before monthlyRent.
    Rewrites 2 conditional-format rules over occCharge after the insert, since Sheets leaves them wrong.
 2. unit › moveOut: emptyValueAllowed → false. Changes the configs only; nothing is sent to the sheet.

== Configs: regenerated from the live sheet ==
Renamed by a person since the last sync; the codebase name follows the header:
  occCharge › rent → monthlyRent      header "Rent" → "Monthly rent"
Names that now point at a different column than at the last sync. Reread src/generated/idNameMap.ts before writing another entry:
  unit › moveIn     was c:unt:k2PqA9x, now c:unt:Rb7sW0e
  unit › moveOut    was c:unt:Rb7sW0e, now c:unt:k2PqA9x
Rewritten in alterations.ts, so each entry still names the column it named before:
  occCharge.newColumns[0].before    "rent" → "monthlyRent"
  unit.columns.moveIn → unit.columns.moveOut    (c:unt:k2PqA9x, which a person retitled "Move out")
Not rewritten: claims and endpoint code. tsc fails where they say rent. A name that moved still type-checks, so code using unit's moveIn or moveOut needs reading by eye.

== Next ==
Nothing held: npm run app:sync -- --send sends this plan.
```

## Your reply

1. Say exactly what you do next, in order: each command, each file edit, each message to the developer (quote it). Then stop.
2. Then say what in the output was unclear, or tempted you toward a step you decided against.
3. Last, separately: was anything about how entries in alterations.ts are named unclear?
