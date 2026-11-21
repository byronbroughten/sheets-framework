import type {
  SyncCheck,
  SyncOutsideReferenceEffect,
} from "../../01_SpreadsheetSchema/configReaders/syncRunResult";
import { reportText } from "../reportText";

const { indent, counted, subjectLabel } = reportText;

export const checkSectionName = "Check before sending";

const cellWidth = 16;

export const checkLines = {
  of(check: SyncCheck): string[] {
    if (check.kind === "outsideReferences") {
      const { subject, effect, references } = check;
      const count = references.length;
      const words = outsideReferenceWords[effect];
      const after = count === 1 ? words.afterOne : words.afterMany;
      return [
        `${subjectLabel(subject)}: ${counted(count, "formula", "formulas")} ${words.where} ${readsVerb(count)} it. After the deletion ${after}:`,
        ...indent(
          2,
          references.map(
            ({ cell, formula }) => `${cell.padEnd(cellWidth)} ${formula}`,
          ),
        ),
      ];
    }
    if (check.kind === "followedSwap") {
      return [
        `${subjectLabel(check.subject)}: sync rewrote this entry across a swap, so it still names the column it named before, now headed "${check.header}". Check that column is the one your task means.`,
      ];
    }
    if (check.kind === "handDrawnRule") {
      const { table, tab, range, formula, handle, cause } = check;
      return [
        `${table}: 1 hand-drawn rule ${cause}: on tab ${tab}, ${range}, ${formula}. Its handle: ${handle}. Fix it by hand in Sheets, or remove it with a conditionalFormats.remove entry naming that handle.`,
      ];
    }
    return [check.text];
  },
  // The plan's line for a deletion that outside formulas read, pointing at its check.
  outsideReferencesEffect(
    count: number,
    effect: SyncOutsideReferenceEffect,
  ): string {
    const formulas = counted(count, "formula", "formulas");
    const after = outsideReferenceWords[effect].inPlan;
    return `${formulas} elsewhere ${readsVerb(count)} it and will ${after} after the deletion (see ${checkSectionName}).`;
  },
};

const outsideReferenceWords: Record<
  SyncOutsideReferenceEffect,
  { where: string; afterOne: string; afterMany: string; inPlan: string }
> = {
  blank: {
    where: "outside this Table",
    afterOne: "it reads blank cells, with no error",
    afterMany: "they read blank cells, with no error",
    inPlan: "read blank cells",
  },
  ref: {
    where: "on another tab",
    afterOne: "it shows #REF!",
    afterMany: "they show #REF!",
    inPlan: "show #REF!",
  },
};

function readsVerb(count: number): string {
  return count === 1 ? "reads" : "read";
}
