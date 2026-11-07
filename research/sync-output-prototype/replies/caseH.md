# Case H: the agent's reply (Sonnet, fresh, one Read)

**1. What I do next, in order**

1. No edit to `src/alterations.ts`. Sync already emptied it and wrote `src/generated/`. I don't touch `src/generated/` and I don't add a `conditionalFormats.remove` entry.
2. Run `npm run check`.
3. If check passes, commit `src/generated/` and `src/alterations.ts` with my change, as the Next block says. Then send the developer the message below.
4. If check fails, I look at where. If every error traces to `household › phone` (TEXT to NUMBER), I stop there. I don't edit household code to fit NUMBER, and I don't change the sheet's column type back. That change is outside my task, and a person made it on purpose or by accident. I don't commit until the developer answers. If any error does not trace to phone, it is mine, and I fix it.
5. No further `--send`, `--confirm` or `--force`. Nothing was held, and none is needed.

Message to the developer (the check line changes to the failing case if tsc fails):

"Done: occPayment's Processed column is now just before Notes, and 'Receipt no.' (TEXT) is added at the Table's end (c:opy:Xw3nB7d). Both alterations landed and alterations.ts is empty. Sync also did three things on its own. It rewrote 3 conditional-format rules after the move. It re-spanned one rule (occPayment·all·=${processed}=FALSE·bg#fff2cc) from row 127 to the last body row, 212. It filled 3 blank IDs on household, rows 140 to 142.

Two things need you:
1. A hand-drawn conditional-format rule on tab Occ Payment, A3:N3 with formula =$H3="", could not be rewritten. The move left it pointing at the wrong column, and it covers a head row. It belongs to the sheet, so I haven't touched it. Fix it by hand in Sheets, or tell me to remove it with a conditionalFormats.remove entry.
2. Someone changed household › phone from TEXT to NUMBER in the live sheet. Regeneration picked it up, so `tsc` fails wherever code relied on TEXT. That is unrelated to this task, so I haven't changed any household code. Do you want me to adapt the code to NUMBER, or is the sheet change a mistake to revert by hand?

Check result: [passed / failed only on household phone]. [Committed / not committed yet.]"

Then I stop.

**2. What was unclear or tempting**

- The Notices heading says "nothing to do unless your task needs it", but the notice itself says my move caused the wrong pointer. I read the heading as "no action by me", and still reported it, because the developer would otherwise see a broken highlight with no explanation.
- I was tempted to add `conditionalFormats.remove` for that rule. I decided against it because the rule is hand-drawn, so its traits belong to the sheet, and deleting it is the developer's call. The notice also doesn't say which column the rule should point at now, or whether the row-3 head-row coverage was intended.
- I was tempted to fix the phone type errors, or hand-edit `src/generated/`. I decided against both. The first is out of scope, and the second is forbidden.
- "Also sent" is unrequested writes: the re-span from 127 to 212 and the filled IDs. They were not held, so no confirm is needed. It isn't clear whether I should ask the developer about them. I only reported them.
- Whether to commit while check is red. The Next block says commit, and commits are free, but I'd rather ask first.

**3. Naming in `alterations.ts`**

Yes, somewhat unclear:
- The output mixes name kinds. The alteration lines use codebase names (`occPayment › processed`). The notice uses the tab's sheet name (`Occ Payment`). The new column appears as the header "Receipt no." plus an ID, and the output never gives its codebase name. I'd have to read `src/generated/idNameMap.ts` to learn it, and I need that name to reference the column in code or in a later alteration entry.
- The rule label `occPayment·all·=${processed}=FALSE·bg#fff2cc` uses a column's codebase name inside a formula. I assume that is also how a `conditionalFormats.remove` entry would identify a rule, but the output doesn't say. It also doesn't say how to identify the hand-drawn rule, which has only a range and a raw `$H3` reference.
