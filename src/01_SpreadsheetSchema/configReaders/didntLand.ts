import {
  type IntendedTrait,
  isColumnDeletion,
  isSameValue,
  isSheetTrait,
  isTableDeletion,
  subjectOf,
  type TraitReadings,
  type TraitValue,
  valueText,
} from "./SyncPlanner";
import type { SyncPlanItem, SyncSendLine, SyncSubject } from "./syncRunResult";

function subjectKey(subject: SyncSubject): string {
  return JSON.stringify(subject);
}

// A filter callback keeping the sent subjects the refetch didn't show Sheets keeping another value for.
export function hasLanded(
  didntLand: SyncSendLine[],
): (subject: SyncSubject) => boolean {
  const didntLandKeys = new Set(
    didntLand.map(({ subject }) => subjectKey(subject)),
  );
  return (subject) => !didntLandKeys.has(subjectKey(subject));
}

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
    plan.map(({ subject, what }) => [subjectKey(subject), what]),
  );
  return sentTraits.filter(isSheetTrait).flatMap((intended): SyncSendLine[] => {
    const kept = keptValue(intended, live);
    if (isSameValue(intended.value, kept)) return [];
    const subject = subjectOf(intended);
    const what = whatBySubject.get(subjectKey(subject));
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
  // A deletion keeps no trait on the sheet: it landed when its Table or column is gone.
  if (isTableDeletion(intended)) return table === undefined;
  if (intended.kind === "table") return table?.traits[intended.trait];
  const column = table?.columns[intended.column];
  if (isColumnDeletion(intended)) return column === undefined;
  return column?.[intended.trait];
}
