import {
  type IntendedTrait,
  isSameValue,
  isSheetTrait,
  subjectOf,
  type TraitReadings,
  type TraitValue,
  valueText,
} from "./SyncPlanner";
import type { SyncPlanItem, SyncSendLine } from "./syncRunResult";

// The sent sheet traits the refetch shows Sheets kept another value for, worded as the plan item plus what stayed.
export function didntLandOf({
  sentTraits,
  plan,
  live,
}: {
  sentTraits: IntendedTrait[];
  plan: SyncPlanItem[];
  live: TraitReadings;
}): SyncSendLine[] {
  const whatBySubject = new Map(
    plan.map(({ subject, what }) => [JSON.stringify(subject), what]),
  );
  return sentTraits.filter(isSheetTrait).flatMap((intended): SyncSendLine[] => {
    const kept = keptValue(intended, live);
    if (isSameValue(intended.value, kept)) return [];
    const subject = subjectOf(intended);
    const what = whatBySubject.get(JSON.stringify(subject));
    if (what === undefined) return [];
    return [{ subject, what: `${what}; Sheets kept ${valueText(kept)}` }];
  });
}

function keptValue(
  intended: IntendedTrait,
  live: TraitReadings,
): TraitValue | undefined {
  if (intended.kind === "tab") return live.tabs[intended.tab]?.[intended.trait];
  const table = live.tables[intended.table];
  if (intended.kind === "table") return table?.traits[intended.trait];
  return table?.columns[intended.column]?.[intended.trait];
}
