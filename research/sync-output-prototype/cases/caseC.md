# Case C

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

On dev, make item's Weight column a Number column, and colour negative weights red.

You added this to `dev/alterations.ts`:

```ts
export const alterations = makeAlterations({
  tables: {
    item: {
      columns: { weight: { columnType: "NUMBER" } },
      conditionalFormats: {
        add: [{ columns: ["weight"], condition: "NUMBER_LESS", values: ["0"], format: { textColor: "#cc0000" } }],
      },
    },
  },
});
```

Then you ran the dry run.

## Terminal

```
$ npm run dev:sync
sync · dev · dry run
READY, 1 HELD: --send sends only with --confirm (see Held). This was a dry run: nothing was sent.

== Held: 1 item, sent only with --confirm ==
item  dropdown style reset (dropdownStyleReset: "ask"):
  Sending item's column list resets these dropdowns to plain chips and clears their option colours: status, category.
  A quirk of Google's API: any column-list send resets them, and the API can neither read nor set them.
Token b21e07: binds this target and exactly this list.
Rerun with --send --confirm b21e07 to send these.

== Plan: what --send sends, in order ==
 1. item: set weight's type to NUMBER (resends all 9 columns).
    Resets 2 dropdowns' style and colours (held, above).
 2. item: add rule item·weight·NUMBER_LESS(0)·text#cc0000.
    Added first; it outranks 1 rule on these cells: item·all·=${status}="Retired"·bg#eeeeee.

== Notices: nothing to do unless your task needs it ==
item: ID is the first column. Sheets' Table name chip covers the cell above it, so item's Fill Row IDs checkbox can't be clicked. Moving ID out of first place fixes it.

== Configs: regenerated from the live sheet ==
Refreshed. No changes.

== Next ==
--send refuses this plan without --confirm b21e07, because of the held item.
```

## Your reply

1. Say exactly what you do next, in order: each command, each file edit, each message to the developer (quote it). Then stop.
2. Then say what in the output was unclear, or tempted you toward a step you decided against.
3. Last, separately: was anything about how entries in alterations.ts are named unclear?
