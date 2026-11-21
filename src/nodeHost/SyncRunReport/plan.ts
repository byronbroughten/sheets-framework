import type {
  SyncPlanEffect,
  SyncPlanItem,
} from "../../01_SpreadsheetSchema/configReaders/syncRunResult";
import { ReportText } from "../ReportText";
import { checkLines } from "./checks";
import { sendLines } from "./sends";

const { indent, formatCount, counted } = ReportText;

export const planLines = {
  of(plan: SyncPlanItem[]): string[] {
    return plan.flatMap((item, index) => [
      `${String(index + 1).padStart(2)}. ${sendLines.of(item)}`,
      ...indent(4, item.effects.map(effectLine)),
    ]);
  },
};

function effectLine(effect: SyncPlanEffect): string {
  if (effect.kind === "dropdownReset") {
    const { columns, setting } = effect;
    const dropdowns = counted(columns.length, "dropdown", "dropdowns");
    const why =
      setting === "ask"
        ? 'dropdownStyleReset: "ask", so held, above'
        : 'dropdownStyleReset: "allow"';
    return `Resets the style and colours of ${dropdowns}: ${columns.join(", ")} (${why}). This is a quirk of Google's API: setting any column's type resends the Table's column list, which resets every dropdown on it.`;
  }
  if (effect.kind === "addedFirst") {
    const { outranks } = effect;
    if (outranks.length === 0) {
      return "Added first; it outranks no rule on these cells.";
    }
    return `Added first; it outranks ${counted(outranks.length, "rule", "rules")} on these cells: ${outranks.join(", ")}.`;
  }
  if (effect.kind === "rulesRewritten" && effect.after === "append") {
    return `Widens ${ruleCount(effect.count)} over ${effect.table} after the append, since Sheets leaves them short of the new columns.`;
  }
  if (effect.kind === "rulesRewritten") {
    return `Rewrites ${ruleCount(effect.count)} over ${effect.table} after the ${effect.after}, since Sheets leaves them wrong.`;
  }
  if (effect.kind === "ruleRespanned") {
    return `Re-spans 1 rule that ended at row ${effect.fromRow} to the Table's last body row, ${effect.toRow}: ${effect.handle}.`;
  }
  if (effect.kind === "rulesRemoved") {
    const { handles } = effect;
    return `Removes ${ruleCount(handles.length)} wholly over it: ${handles.join(", ")}.`;
  }
  if (effect.kind === "rulePartlyOver") {
    return `Leaves 1 rule partly over it, on ${effect.tab} ${effect.range} (${effect.formula}); Sheets keeps it over the blank cells.`;
  }
  if (effect.kind === "outsideReferences") {
    return checkLines.outsideReferencesEffect(effect.count, effect.effect);
  }
  return `Overwrites ${counted(effect.cells, "cell", "cells")} of entered data (held, above); the other ${formatCount(effect.blank)} are blank.`;
}

function ruleCount(count: number): string {
  return counted(count, "conditional-format rule", "conditional-format rules");
}
