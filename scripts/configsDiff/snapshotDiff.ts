import { Obj } from "@byronbroughten/utils/obj";
import { Val } from "@byronbroughten/utils/val";

import type {
  FixedList,
  FixedListMember,
  ValueConfigBase,
} from "../../src/01_SpreadsheetSchema/configReaders/makeConfigs.ts";
import {
  type ConfigEntry,
  type ConfigSnapshot,
  type IdentifiedEntry,
  type IdField,
  type RuleSnapshot,
  type TabEntry,
  valueKindLabels,
} from "../configsDiff.ts";

interface BeforeAfter<SV> {
  before: SV;
  after: SV;
}

interface Matched<EN extends object> extends BeforeAfter<EN> {
  beforeKey: string;
  afterKey: string;
}

type Pairing<EN extends object> =
  | ({ kind: "matched" } & Matched<EN>)
  | { kind: "removed"; key: string; entry: EN }
  | { kind: "added"; key: string; entry: EN };

const ownerWords = { tables: "Table", tabs: "Tab" } as const;

export interface Section {
  header: string;
  changes: string[];
  lines: string[];
}

export const snapshotDiff = {
  sections({ before, after }: BeforeAfter<ConfigSnapshot>): Section[] {
    return [
      ...diffTables({ before, after }),
      ...diffTabs({ before: before.tabConfigs, after: after.tabConfigs }),
      ...diffRetiredIdPrefixes({
        before: before.retiredIdPrefixes,
        after: after.retiredIdPrefixes,
      }),
      ...diffValues({ before: before.valueConfigs, after: after.valueConfigs }),
      ...diffRules({
        before: before.conditionalFormatConfigs,
        after: after.conditionalFormatConfigs,
      }),
    ];
  },
  sectionHeader({ header, changes }: Section): string {
    if (changes.length === 0) return `${header}:`;
    return `${header}: ${changes.join("; ")}`;
  },
};

function diffTables({ before, after }: BeforeAfter<ConfigSnapshot>): Section[] {
  const pairings = pairEntries(
    { before: before.tableConfigs, after: after.tableConfigs },
    (_key, { tableId }) => tableId,
  );
  return pairings.flatMap((pairing) => {
    if (pairing.kind === "removed") {
      const columns = before.columnConfigs[pairing.key] ?? {};
      const lines = diffCols({ before: columns, after: {} });
      return [tableSection(pairing.key, ["removed"], lines)];
    }
    if (pairing.kind === "added") {
      const columns = after.columnConfigs[pairing.key] ?? {};
      const lines = diffCols({ before: {}, after: columns });
      return [tableSection(pairing.key, ["added"], lines)];
    }
    const changes = entryChanges(pairing, "tableId");
    const lines = diffCols({
      before: before.columnConfigs[pairing.beforeKey] ?? {},
      after: after.columnConfigs[pairing.afterKey] ?? {},
    });
    if (changes.length + lines.length === 0) return [];
    return [tableSection(renameLabel(pairing), changes, lines)];
  });
}

function tableSection(
  label: string,
  changes: string[],
  lines: string[],
): Section {
  return { header: `${ownerWords.tables} ${label}`, changes, lines };
}

function diffCols(
  columns: BeforeAfter<Record<string, IdentifiedEntry<"columnId">>>,
): string[] {
  const pairings = pairEntries(columns, (_key, { columnId }) => columnId);
  return pairings.flatMap((pairing) => {
    if (pairing.kind !== "matched") {
      return [colLine(pairing.key, [pairing.kind])];
    }
    const changes = entryChanges(pairing, "columnId");
    if (changes.length === 0) return [];
    return [colLine(renameLabel(pairing), changes)];
  });
}

function colLine(label: string, changes: string[]): string {
  return `  column ${label}: ${changes.join("; ")}`;
}

function diffTabs(tabs: BeforeAfter<Record<string, TabEntry>>): Section[] {
  const pairings = pairEntries(tabs, (_key, { sheetGid }) => sheetGid);
  return pairings.flatMap((pairing) => {
    if (pairing.kind !== "matched") {
      return [tabSection(pairing.key, [pairing.kind])];
    }
    const changes = entryChanges(pairing, "sheetGid");
    if (changes.length === 0) return [];
    return [tabSection(renameLabel(pairing), changes)];
  });
}

function tabSection(label: string, changes: string[]): Section {
  return { header: `${ownerWords.tabs} ${label}`, changes, lines: [] };
}

function diffRetiredIdPrefixes(prefixes: BeforeAfter<string[]>): Section[] {
  const changes = memberChanges(prefixes);
  if (changes.length === 0) return [];
  return [{ header: "Retired ID prefixes", changes, lines: [] }];
}

// Removed and matched entries keep the before order; added entries follow in the after order.
function pairEntries<EN extends object>(
  { before, after }: BeforeAfter<Record<string, EN>>,
  identity: (key: string, entry: EN) => unknown,
): Pairing<EN>[] {
  const afterKeyById = new Map(
    Object.entries(after).map(([key, entry]) => [identity(key, entry), key]),
  );
  const beforeIds = new Set(
    Object.entries(before).map(([key, entry]) => identity(key, entry)),
  );
  const matchedAndRemoved = Object.entries(before).map(
    ([beforeKey, entry]): Pairing<EN> => {
      const afterKey = afterKeyById.get(identity(beforeKey, entry));
      if (afterKey === undefined) {
        return { kind: "removed", key: beforeKey, entry };
      }
      return {
        kind: "matched",
        beforeKey,
        afterKey,
        before: entry,
        after: Val.assert(after[afterKey], "after config entry"),
      };
    },
  );
  const added = Object.entries(after)
    .filter(([key, entry]) => !beforeIds.has(identity(key, entry)))
    .map(([key, entry]): Pairing<EN> => ({ kind: "added", key, entry }));
  return [...matchedAndRemoved, ...added];
}

function entryChanges(
  { beforeKey, afterKey, before, after }: Matched<ConfigEntry>,
  idField: IdField | "sheetGid",
): string[] {
  const renamed = beforeKey === afterKey ? [] : ["renamed"];
  const changes = fieldChanges({ before, after }).filter(
    ({ field }) => field !== idField,
  );
  return [...renamed, ...changes.map(({ change }) => change)];
}

function fieldChanges({
  before,
  after,
}: BeforeAfter<ConfigEntry>): { field: string; change: string }[] {
  const fields = [...new Set([...Object.keys(before), ...Object.keys(after)])];
  return fields.flatMap((field) => {
    const old = JSON.stringify(before[field]);
    const next = JSON.stringify(after[field]);
    if (old === next) return [];
    return [{ field, change: `${field} ${old} → ${next}` }];
  });
}

function renameLabel({ beforeKey, afterKey }: Matched<object>): string {
  return beforeKey === afterKey ? beforeKey : `${beforeKey} → ${afterKey}`;
}

function diffValues(
  values: BeforeAfter<Record<string, ValueConfigBase>>,
): Section[] {
  return pairEntries(values, (name) => name).flatMap((pairing) => {
    if (pairing.kind !== "matched") {
      return [valueSection(pairing.key, [pairing.kind])];
    }
    const changes = valueConfigChanges(pairing);
    if (changes.length === 0) return [];
    return [valueSection(pairing.beforeKey, changes)];
  });
}

function valueConfigChanges({
  before,
  after,
}: BeforeAfter<ValueConfigBase>): string[] {
  if (isFixedList(before) && isFixedList(after)) {
    return memberChanges({ before, after });
  }
  const beforeKind = valueKindLabel(before);
  const afterKind = valueKindLabel(after);
  return beforeKind === afterKind ? [] : [`${beforeKind} → ${afterKind}`];
}

function isFixedList(config: ValueConfigBase): config is FixedList {
  return Array.isArray(config);
}

function valueKindLabel(config: ValueConfigBase): string {
  if (isFixedList(config)) return valueKindLabels.fixedList;
  return valueKindLabels.liveReference;
}

function memberChanges({ before, after }: BeforeAfter<FixedList>): string[] {
  const added = after.filter((member) => !before.includes(member));
  const removed = before.filter((member) => !after.includes(member));
  const changes = [
    ...memberListChange("added", added),
    ...memberListChange("removed", removed),
  ];
  if (
    changes.length === 0 &&
    JSON.stringify(before) !== JSON.stringify(after)
  ) {
    return ["reordered"];
  }
  return changes;
}

function memberListChange(verb: string, members: FixedListMember[]): string[] {
  if (members.length === 0) return [];
  const quoted = members.map((member) => JSON.stringify(member));
  return [`${verb} ${quoted.join(", ")}`];
}

function valueSection(name: string, changes: string[]): Section {
  return { header: `Value ${name}`, changes, lines: [] };
}

// Rules pair by handle, which is built from their content; groups pair by Table or tab key, so a renamed one reads as removed and added.
function diffRules({ before, after }: BeforeAfter<RuleSnapshot>): Section[] {
  return Obj.keys(ownerWords).flatMap((owner) => {
    const groups = { before: before[owner], after: after[owner] };
    return pairEntries(groups, (key) => key).flatMap((pairing) => {
      const lines = diffRuleGroup(groupPair(pairing));
      if (lines.length === 0) return [];
      const label =
        pairing.kind === "matched" ? pairing.beforeKey : pairing.key;
      return [ruleSection(`${ownerWords[owner]} ${label}`, lines)];
    });
  });
}

function ruleSection(ownerLabel: string, lines: string[]): Section {
  return { header: `${ownerLabel} conditional formats`, changes: [], lines };
}

function groupPair<EN extends object>(
  pairing: Pairing<Record<string, EN>>,
): BeforeAfter<Record<string, EN>> {
  if (pairing.kind === "removed") return { before: pairing.entry, after: {} };
  if (pairing.kind === "added") return { before: {}, after: pairing.entry };
  return pairing;
}

function diffRuleGroup(
  rules: BeforeAfter<Record<string, ConfigEntry>>,
): string[] {
  return pairEntries(rules, (handle) => handle).flatMap((pairing) => {
    if (pairing.kind !== "matched") {
      return [ruleLine(pairing.key, [pairing.kind])];
    }
    const changes = fieldChanges(pairing).map(({ change }) => change);
    if (changes.length === 0) return [];
    return [ruleLine(pairing.beforeKey, changes)];
  });
}

function ruleLine(handle: string, changes: string[]): string {
  return `  rule ${handle}: ${changes.join("; ")}`;
}
