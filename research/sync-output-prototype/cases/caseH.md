# Case H

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

Move occPayment's Processed column to just before Notes, and add a "Receipt no." Text column at the end.

You added this to `src/alterations.ts`:

```ts
export const alterations = makeAlterations({
  tables: {
    occPayment: {
      columns: { processed: { before: "notes" } },
      newColumns: [{ header: "Receipt no.", columnType: "TEXT" }],
    },
  },
});
```

Your dry run held nothing, so you ran the send.

## Terminal

```
$ npm run app:sync -- --send
sync · app · send
SENT: 2 alterations landed; 3 blank IDs filled.

== Landed: removed from alterations.ts ==
occPayment › processed: moved before notes.
occPayment: created column "Receipt no." as c:opy:Xw3nB7d, TEXT, at the Table's end.
alterations.ts is now empty.

== Also sent ==
 1. occPayment: rewrote 3 conditional-format rules after the move, since Sheets leaves them wrong.
 2. occPayment: re-spanned 1 rule that ended at row 127 to the Table's last body row, 212: occPayment·all·=${processed}=FALSE·bg#fff2cc.
 3. household: filled 3 blank IDs: hh-8Kd2Lq0, hh-T41xZpa, hh-c9WmE2r (rows 140 to 142).

== Notices: nothing to do unless your task needs it ==
occPayment: 1 hand-drawn rule couldn't be rewritten, and the move left it pointing at the wrong column: on tab Occ Payment, A3:N3 (it covers a head row), =$H3="". Fix it by hand in Sheets, or remove it with a conditionalFormats.remove entry.

== Configs: regenerated from the live sheet ==
Typed traits a person changed since the last sync, on Tables with API access. tsc fails where code relied on the old ones:
  household › phone    columnType TEXT → NUMBER

== Next ==
sync wrote src/generated/ and src/alterations.ts. Run npm run check, then commit them with your change.
```

## Your reply

1. Say exactly what you do next, in order: each command, each file edit, each message to the developer (quote it). Then stop.
2. Then say what in the output was unclear, or tempted you toward a step you decided against.
3. Last, separately: was anything about how entries in alterations.ts are named unclear?
