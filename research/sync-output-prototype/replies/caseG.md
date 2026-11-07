# Case G: the agent's reply (Sonnet, fresh, one Read)

**1. What I do next, in order**

1. Run the dry run, as the output's Next section says: `npm run app:sync`. It sends nothing and regenerates `src/generated/` from the live sheet.
2. I make no file edits yet. The "Late fee" newColumns entry was already removed from `alterations.ts` when it landed. The notes formula entry is still there and stays as it is.
3. I read the new dry run. Expected result: the notes formula alteration is planned again, and it either lists held items or doesn't.
   - If it lists held items, I paste the held block to the developer with a "why" for each item. I don't reuse `91be4f`. The sheet changed after the first dry run, so the held set may be different (for example more or fewer than 12 cells, or different row IDs). The developer's "yes, those 12 can go" covers only the 12 cells I showed them.
   - I message the developer: "The send stopped partway. The 'Late fee' Currency column was created on occCharge (flush 1). The notes formula change was refused because the sheet changed after my dry run, so your yes for the 12 cells doesn't cover what's held now. The new dry run holds: [pasted held block]. Why each is held: [reason per item]. Which of these are you OK with losing? I'll send only if every item is one you agreed to."
   - I run `npm run app:sync -- --send --confirm <new token>` only if every held item is among the 12 row IDs the developer approved. Because it's an `app:*` command with `--confirm`, I wait for the developer's permission to run it. If anything differs from the 12 approved cells, I stop and discuss instead.
   - If the dry run lists no held items, I run `npm run app:sync -- --send` with no `--confirm`.
4. After the notes alteration lands, I run `npm run check`, then report to the developer.

**2. What was unclear or tempting**

- I was tempted to rerun the same command and just reuse `--confirm 91be4f`. I decided against it because the output says the token no longer covers the held items.
- I was tempted to treat the developer's earlier yes as still valid. The sheet changed after the dry run, and I can't tell whether the 12 cells are still the same 12.
- I was tempted to edit `alterations.ts` to remove or re-add entries, or to edit `src/generated/`. The output says the landed entry is already removed, and generated files are never edited by hand. Nothing needed changing.
- The output doesn't say what changed in the sheet. It doesn't say whether the sheet change was in the notes column or somewhere else, so I can't judge from this output whether the old approval still applies.
- It doesn't say whether the unsent formula alteration will still be correct against the regenerated configs. The formula refers to `occCharge[Late fee]`, and the column now exists with the generated key `lateFee`. I'd check that the formula's reference still reads correctly in the dry run.
- "Refreshed from the sheet after flush 1" confirms the configs are up to date, but I'd still rely on the dry run to confirm that.

**3. Naming of entries in alterations.ts**

Yes, one thing was unclear. The rules say entries are keyed by codebase name, as listed in `src/generated/idNameMap.ts`. My entry used `columns: { notes: ... }` for an existing column, while the new column was given by header text ("Late fee"). The output shows the new column as `lateFee` in the generated configs and as the ID `c:och:Q8x1mZp`, and it shows the held cells as "occCharge › notes". I'm fairly confident `notes` is the codebase name, but the rules never say whether `columns:` keys are codebase names, headers or IDs. I'd check `idNameMap.ts` after the dry run. Also unclear: the formula uses the sheet header `occCharge[Late fee]`, not the codebase name `lateFee`. The rules don't say which kind of name a formula should use.
