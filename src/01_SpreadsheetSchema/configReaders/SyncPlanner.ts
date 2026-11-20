import type { DropdownStyleReset } from "../../00_Source/RawSource/RawSource";
import type { ColumnTypeBaseline } from "../columnBaseline";
import { SpreadsheetBaseSchema } from "../SpreadsheetBaseSchema";
import { columnNameOfHeader } from "./columnConfigsTypes";
import type { PendingAlterations } from "./intendedState";
import { DeletionRefuser } from "./SyncPlanner/DeletionRefuser";
import {
  entriesOf,
  type WideAlterations,
  type WideNewTable,
  type WideRule,
} from "./SyncPlanner/wideAlterations";
import type {
  ForcedTrait,
  SyncConflict,
  SyncHeldBackTrait,
  SyncPlanEffect,
  SyncPlanItem,
  SyncRefusal,
  SyncRunResult,
  SyncSendLine,
  SyncSheetStop,
  SyncSubject,
} from "./syncRunResult";

export type TraitValue = string | number | boolean | ColumnTypeBaseline;
type TraitRecord = Record<string, TraitValue | undefined>;

export interface TableReading {
  traits: TraitRecord;
  columns: Record<string, TraitRecord | undefined>;
}

// One side of the comparison: the previous configs or the live sheet, by codebase name.
export interface TraitReadings {
  tables: Record<string, TableReading | undefined>;
  tabs: Record<string, TraitRecord | undefined>;
}

// Each Table's tab, both by codebase name.
export type TabsByTable = Record<string, string | undefined>;

export interface SyncPlannerProps {
  alterations: PendingAlterations;
  tabsByTable: TabsByTable;
  // Only ever this run's --force arguments, so a yes can't outlive its run.
  forcedTraits: ForcedTrait[];
  // The app's setting, which a column type's dropdownReset effect names.
  dropdownStyleReset: DropdownStyleReset;
}

export interface PlanReadings {
  previous: TraitReadings;
  live: TraitReadings;
}

export interface PlanProps extends PlanReadings {
  // Stops found reading the sheet, outside the comparison, that still stop every send.
  sheetStops?: SyncSheetStop[];
}

// One `newColumns` entry, by its Table's codebase name; `column` is the name the regen gives its header.
export interface PlannedNewColumn {
  table: string;
  // The entry's path in its Table entry, as `newColumns[0]`.
  path: string;
  column: string;
  header: string;
  before?: string;
  emptyValueAllowed: boolean;
}

// A tab the run adds, or one already there by its codebase name, with `gap` empty columns right of its rightmost Table.
export type PlannedNewTablePlacement =
  | { kind: "newTab"; title: string }
  | { kind: "besideTables"; tab: string; gap: 0 | 1 };

// A new Table's column, named by the name the regen gives its header.
export interface PlannedNewTableColumn {
  header: string;
  column: string;
  emptyValueAllowed: boolean;
  // In brace form, as the alterations give them.
  formula?: string;
  columnType?: ColumnTypeBaseline;
}

// One `newTabs` or `newTables` Table; `table` is the codebase name the regen gives its name.
export interface PlannedNewTable {
  table: string;
  name: string;
  placement: PlannedNewTablePlacement;
  columns: PlannedNewTableColumn[];
  // Its `conditionalFormats.add`, compiled once the Table's created.
  rules: WideRule[];
}

// One `conditionalFormats.remove` handle, by its Table's codebase name.
export interface PlannedRuleRemoval {
  table: string;
  handle: string;
}

// One `conditionalFormats.replace` entry, by its Table's codebase name: the handle it replaces and the rule it writes there.
export interface PlannedRuleReplacement {
  table: string;
  handle: string;
  rule: WideRule;
}

// One `conditionalFormats.add` entry, by its Table's codebase name and its place in the list.
export interface PlannedRuleAddition {
  table: string;
  position: number;
  rule: WideRule;
}

// A rule ending short of its Table's last body row, by its Table's codebase name; rows are 1-based, as a person reads them.
export interface PlannedRuleRespan {
  table: string;
  handle: string;
  fromRow: number;
  toRow: number;
}

export interface SyncPlannerResult extends Pick<
  SyncRunResult,
  "checks" | "stops" | "plan"
> {
  isStopped: boolean;
  configs: { heldBack: SyncHeldBackTrait[] };
  // Handed to the flush classifier, which makes each one a held item.
  forcedTraits: ForcedTrait[];
  // What the regenerated configs record if the plan's sends land.
  configTraits: TraitReadings;
  // The plan's sends by address, for the run to queue; empty when stopped.
  sentTraits: IntendedTrait[];
  // The columns the plan creates, for the run to queue; empty when stopped.
  newColumns: PlannedNewColumn[];
  // The Tables the plan creates, for the run to queue; empty when stopped.
  newTables: PlannedNewTable[];
  // The rules the plan removes, for the run to queue; empty when stopped.
  ruleRemovals: PlannedRuleRemoval[];
  // The rules the plan replaces, for the run to queue; empty when stopped.
  ruleReplacements: PlannedRuleReplacement[];
  // The rules the plan adds, for the run to queue; empty when stopped.
  ruleAdditions: PlannedRuleAddition[];
}

interface TableTraitAddress {
  kind: "table";
  table: string;
  trait: string;
}

interface ColumnTraitAddress {
  kind: "column";
  table: string;
  column: string;
  trait: string;
}

interface TabTraitAddress {
  kind: "tab";
  tab: string;
  trait: string;
}

type TraitAddress = TableTraitAddress | ColumnTraitAddress | TabTraitAddress;

interface IntendedValue {
  value: TraitValue;
}

export type IntendedTrait = TraitAddress & IntendedValue;

// The traits a tab entry sets.
export const tabTraits = [
  "title",
  "color",
  "before",
  "frozenRows",
  "frozenColumns",
  "hidden",
] as const;

// The sent traits the run refetches after sending, so the refetch can show whether each landed; a tab's other traits land in the working view.
export const sheetTraits = {
  table: new Set<string>(["name", "deleted"]),
  column: new Set<string>([
    "header",
    "columnType",
    "formula",
    "before",
    "deleted",
  ]),
  tab: new Set<string>(["deleted"]),
} as const;

// A deletion's one value, `true`, says nothing a plan line could show.
const deletionWhat = {
  table: "delete the Table",
  column: "delete the column",
  tab: "delete the tab",
} as const;

// How a `newColumns` entry reads on the plan, landed and didntLand lists.
export const newColumnLines = {
  // Every entry's subject path; landing one drops the whole list from alterations.ts.
  path: "newColumns",
  // Where a new column goes, as its plan line and its created line say it.
  position({ before }: PlannedNewColumn): string {
    if (before === undefined) return "at the Table's end";
    return `before ${before}`;
  },
  // Its plan line, and its didntLand line when the refetch can't find it.
  sendLine(newColumn: PlannedNewColumn): SyncSendLine {
    return {
      subject: { table: newColumn.table, path: newColumnLines.path },
      what: `create "${newColumn.header}" ${newColumnLines.position(newColumn)}`,
    };
  },
  isSubject(subject: SyncSubject): boolean {
    return "table" in subject && subject.path === newColumnLines.path;
  },
};

// How a `newTabs` or `newTables` entry reads on the plan, landed and didntLand lists.
export const newTableLines = {
  // A new tab's own line, whose landing drops its `newTabs` entry with its Tables'.
  tabPath: "newTabs",
  // Each Table's subject path on a new tab, apart from the tab's own line.
  tableOnNewTabPath: "newTabs.tables",
  // Every Table's subject path on a tab already there; landing one drops the tab's whole list.
  besidePath: "newTables",
  tabLine(title: string): SyncSendLine {
    return {
      subject: { tab: title, path: newTableLines.tabPath },
      what: "create the tab",
    };
  },
  // Its plan line, and its didntLand line when the refetch can't find it.
  sendLine({ name, placement }: PlannedNewTable): SyncSendLine {
    const what = `create Table "${name}"`;
    if (placement.kind === "newTab") {
      return {
        subject: {
          tab: placement.title,
          path: newTableLines.tableOnNewTabPath,
        },
        what,
      };
    }
    const gapText = placement.gap === 0 ? "" : ", one empty column between";
    return {
      subject: { tab: placement.tab, path: newTableLines.besidePath },
      what: `${what} right of the tab's Tables${gapText}`,
    };
  },
  isSubject(subject: SyncSubject): boolean {
    if (!("tab" in subject)) return false;
    return (
      subject.path === newTableLines.tableOnNewTabPath ||
      subject.path === newTableLines.besidePath
    );
  },
};

// How a `conditionalFormats.remove` handle reads on the plan, stop, landed and didntLand lists.
export const ruleRemovalLines = {
  path({ handle }: PlannedRuleRemoval): string {
    return `conditionalFormats.remove["${handle}"]`;
  },
  sendLine(removal: PlannedRuleRemoval): SyncSendLine {
    return {
      subject: { table: removal.table, path: ruleRemovalLines.path(removal) },
      what: "remove the rule",
    };
  },
};

// How a `conditionalFormats.replace` entry reads on the plan, stop, landed and didntLand lists.
export const ruleReplacementLines = {
  path({ handle }: Pick<PlannedRuleReplacement, "handle">): string {
    return `conditionalFormats.replace["${handle}"]`;
  },
  sendLine(replacement: PlannedRuleReplacement): SyncSendLine {
    return {
      subject: {
        table: replacement.table,
        path: ruleReplacementLines.path(replacement),
      },
      what: "replace the rule",
    };
  },
};

// How a `conditionalFormats.add` entry reads on the plan, landed and didntLand lists.
export const ruleAdditionLines = {
  path({ position }: Pick<PlannedRuleAddition, "position">): string {
    return `conditionalFormats.add[${position}]`;
  },
  sendLine(addition: PlannedRuleAddition): SyncSendLine {
    return {
      subject: {
        table: addition.table,
        path: ruleAdditionLines.path(addition),
      },
      what: "add the rule",
    };
  },
};

// How a rule sync stretches to its Table's last body row reads on the plan, landed and didntLand lists; no alterations.ts entry has its path.
export const ruleRespanLines = {
  path({ handle }: Pick<PlannedRuleRespan, "handle">): string {
    return `conditionalFormats["${handle}"]`;
  },
  sendLine(respan: PlannedRuleRespan): SyncSendLine {
    return {
      subject: { table: respan.table, path: ruleRespanLines.path(respan) },
      what: "re-span the rule",
    };
  },
  planItem(respan: PlannedRuleRespan): SyncPlanItem {
    const { handle, fromRow, toRow } = respan;
    return {
      ...ruleRespanLines.sendLine(respan),
      effects: [{ kind: "ruleRespanned", handle, fromRow, toRow }],
    };
  },
};

export function isSheetTrait(intended: IntendedTrait): boolean {
  return sheetTraits[intended.kind].has(intended.trait);
}

interface HeldBackValue {
  intended: IntendedTrait;
  agreed: TraitValue;
}

interface StopContext {
  intended: IntendedTrait;
  agreed: TraitValue | undefined;
  live: TraitValue | undefined;
  label: string;
  cause: string;
  isColumnGone: boolean;
}

const stopWords = {
  personChanged: {
    label: "a person changed this",
    cause: "A person changed this trait on the sheet after the last sync.",
  },
  bothChanged: {
    label: "both changed",
    cause: "You and a person each changed this trait, to different values.",
  },
  columnDeleted: {
    label: "a person deleted this column",
    cause: "The column is in the configs but no longer on the sheet.",
  },
} as const;

const absentText = "(none)";

// Pure: the three-way comparison of intended state against the previous configs and the live sheet.
export class SyncPlanner {
  private alterations: WideAlterations;
  private tabsByTable: TabsByTable;
  private forcedTraits: ForcedTrait[];
  private dropdownStyleReset: DropdownStyleReset;
  constructor(props: SyncPlannerProps) {
    this.alterations = props.alterations;
    this.tabsByTable = props.tabsByTable;
    this.forcedTraits = props.forcedTraits;
    this.dropdownStyleReset = props.dropdownStyleReset;
  }
  static init(props: SyncPlannerProps): SyncPlanner {
    return new SyncPlanner(props);
  }
  private get _deletionRefuser(): DeletionRefuser {
    return new DeletionRefuser({
      alterations: this.alterations,
      tabsByTable: this.tabsByTable,
    });
  }
  // Takes only intended state, so the run can refuse before reading anything.
  get refusals(): SyncRefusal[] {
    return this._deletionRefuser.refusals;
  }
  // From intended state alone, so the run can check each header against the sheet before planning.
  get newColumns(): PlannedNewColumn[] {
    return entriesOf(this.alterations.tables).flatMap(([table, entry]) =>
      (entry.newColumns ?? []).flatMap(
        ({ header, before, emptyValueAllowed }, index): PlannedNewColumn[] => {
          if (header === undefined) return [];
          return [
            {
              table,
              path: `newColumns[${index}]`,
              column: columnNameOfHeader(header),
              header,
              ...(before === undefined ? {} : { before }),
              emptyValueAllowed: emptyValueAllowed ?? false,
            },
          ];
        },
      ),
    );
  }
  // From intended state alone, so the run can check each name and title against the sheet before planning.
  get newTables(): PlannedNewTable[] {
    const onNewTabs = (this.alterations.newTabs ?? []).flatMap(
      ({ title, tables }) =>
        tables.map((newTable) =>
          plannedNewTable(newTable, { kind: "newTab", title }),
        ),
    );
    const besideTables = entriesOf(this.alterations.tabs).flatMap(
      ([tab, { newTables = [] }]) =>
        newTables.map((newTable) =>
          plannedNewTable(newTable, {
            kind: "besideTables",
            tab,
            gap: newTable.gap ?? 0,
          }),
        ),
    );
    return [...onNewTabs, ...besideTables];
  }
  // From intended state alone, so the run can match each handle against the sheet's rules before planning.
  get ruleRemovals(): PlannedRuleRemoval[] {
    // A handle listed twice is one removal: a second delete at its index would take whichever rule moved into it.
    return entriesOf(this.alterations.tables).flatMap(([table, entry]) =>
      [...new Set(entry.conditionalFormats?.remove)].map((handle) => ({
        table,
        handle,
      })),
    );
  }
  // From intended state alone, so the run can match each handle against the sheet's rules before planning.
  get ruleReplacements(): PlannedRuleReplacement[] {
    return entriesOf(this.alterations.tables).flatMap(([table, entry]) =>
      entriesOf(entry.conditionalFormats?.replace).map(([handle, rule]) => ({
        table,
        handle,
        rule,
      })),
    );
  }
  get ruleAdditions(): PlannedRuleAddition[] {
    return entriesOf(this.alterations.tables).flatMap(([table, entry]) =>
      (entry.conditionalFormats?.add ?? []).map((rule, position) => ({
        table,
        position,
        rule,
      })),
    );
  }
  plan({ previous, live, sheetStops = [] }: PlanProps): SyncPlannerResult {
    const stops: SyncConflict[] = [];
    const sends: SyncPlanItem[] = [];
    const forced: ForcedTrait[] = [];
    const sentTraits: IntendedTrait[] = [];
    const heldBack: HeldBackValue[] = [];
    this._alterationDeclarations.forEach((intended) => {
      const sides = comparedSides(intended, previous, live);
      if (sides === undefined) return;
      const effects = this._dropdownResetEffects(intended, live);
      function send(what: string): void {
        sends.push({ subject: subjectOf(intended), what, effects });
        sentTraits.push(intended);
      }
      if (isOrderTrait(intended)) {
        if (!isSameValue(intended.value, sides.live)) {
          send(`set to ${valueText(intended.value)}`);
        }
        return;
      }
      // comparedSides already left out a Table a person deleted, so a live Table is always deleted.
      if (isTableDeletion(intended)) {
        send(deletionWhat.table);
        return;
      }
      // Likewise a tab a person deleted.
      if (isTabDeletion(intended)) {
        send(deletionWhat.tab);
        return;
      }
      const stop = stopOf(intended, sides);
      if (stop === undefined) {
        if (!sendsIntended(intended.value, sides)) return;
        if (isColumnDeletion(intended)) {
          send(deletionWhat.column);
        } else {
          send(`set to ${valueText(intended.value)}`);
        }
        return;
      }
      const forcedTrait = sides.isColumnGone
        ? undefined
        : this._forcedTraitFor(intended);
      if (forcedTrait === undefined) {
        stops.push(conflictOf(stop));
      } else {
        forced.push(forcedTrait);
        send(
          `set to ${valueText(intended.value)}, over the sheet's ${valueText(sides.live)}`,
        );
      }
      // Set for a forced trait too: a stop elsewhere means it isn't sent, and the next run must still see it.
      if (stop.agreed !== undefined) {
        heldBack.push({ intended, agreed: stop.agreed });
      }
    });
    // A Table a person deleted has nowhere to put one.
    const newColumns = onLiveTables(this.newColumns, live);
    sends.push(...effectlessSends(newColumns, newColumnLines));
    // A tab a person deleted has nowhere to put one.
    const newTables = this.newTables.filter(
      ({ placement }) =>
        placement.kind === "newTab" || live.tabs[placement.tab] !== undefined,
    );
    sends.push(...newTableSends(newTables));
    // A Table a person deleted took its rules with it.
    const ruleRemovals = onLiveTables(this.ruleRemovals, live);
    sends.push(...effectlessSends(ruleRemovals, ruleRemovalLines));
    const ruleReplacements = onLiveTables(this.ruleReplacements, live);
    sends.push(...effectlessSends(ruleReplacements, ruleReplacementLines));
    const ruleAdditions = onLiveTables(this.ruleAdditions, live);
    sends.push(...effectlessSends(ruleAdditions, ruleAdditionLines));
    const isStopped = stops.length + sheetStops.length > 0;
    const sentIfUnstopped = isStopped ? [] : sentTraits;
    const newColumnsIfUnstopped = isStopped ? [] : newColumns;
    const newTablesIfUnstopped = isStopped ? [] : newTables;
    const forcedNotHeld = new Set(
      forced.map((trait) => addressKey({ kind: "column", ...trait })),
    );
    const heldBackTraits = heldBack.filter(
      ({ intended }) => isStopped || !forcedNotHeld.has(addressKey(intended)),
    );
    return {
      // The planner can't read what a send breaks outside its subject; the run fills these in from its live read.
      checks: [],
      stops: [...sheetStops, ...stops],
      isStopped,
      plan: sends,
      configs: {
        heldBack: heldBackTraits.map(({ intended, agreed }) => ({
          subject: subjectOf(intended),
          value: valueText(agreed),
        })),
      },
      forcedTraits: forced,
      configTraits: configTraitsOf({
        live,
        sent: sentIfUnstopped,
        heldBack: heldBackTraits,
        newColumns: newColumnsIfUnstopped,
        newTables: newTablesIfUnstopped,
        tabsByTable: this.tabsByTable,
      }),
      sentTraits: sentIfUnstopped,
      newColumns: newColumnsIfUnstopped,
      newTables: newTablesIfUnstopped,
      ruleRemovals: isStopped ? [] : ruleRemovals,
      ruleReplacements: isStopped ? [] : ruleReplacements,
      ruleAdditions: isStopped ? [] : ruleAdditions,
    };
  }
  // Google resets every dropdown on a Table its column list is sent to, a newly set one included: docs/sheets-behaviour.md#creating-and-updating-tables.
  private _dropdownResetEffects(
    intended: IntendedTrait,
    live: TraitReadings,
  ): SyncPlanEffect[] {
    if (intended.kind !== "column" || intended.trait !== "columnType") {
      return [];
    }
    const columns = entriesOf(live.tables[intended.table]?.columns).flatMap(
      ([column, traits]) => {
        if (hasRule(traits.columnType)) return [column];
        if (column === intended.column && hasRule(intended.value)) {
          return [column];
        }
        return [];
      },
    );
    if (columns.length === 0) return [];
    return [
      { kind: "dropdownReset", columns, setting: this.dropdownStyleReset },
    ];
  }
  private _forcedTraitFor(intended: IntendedTrait): ForcedTrait | undefined {
    const key = addressKey(intended);
    return this.forcedTraits.find(
      (trait) => addressKey({ kind: "column", ...trait }) === key,
    );
  }
  private get _alterationDeclarations(): IntendedTrait[] {
    const tableTraits = entriesOf(this.alterations.tables).flatMap(
      ([table, entry]): IntendedTrait[] => [
        ...pickTraits(entry, ["name", "letApiAccess", "deleted"]).map(
          (trait) => ({
            kind: "table" as const,
            table,
            ...trait,
          }),
        ),
        ...entriesOf(entry.columns).flatMap(([column, columnEntry]) =>
          pickTraits(columnEntry, [
            "header",
            "columnType",
            "formula",
            "before",
            "emptyValueAllowed",
            "deleted",
          ]).map((trait) => ({
            kind: "column" as const,
            table,
            column,
            ...trait,
          })),
        ),
      ],
    );
    const tabDeclarations = entriesOf(this.alterations.tabs).flatMap(
      ([tab, entry]): IntendedTrait[] =>
        pickTraits(entry, [...tabTraits, "deleted"]).map((trait) => ({
          kind: "tab",
          tab,
          ...trait,
        })),
    );
    return [...tableTraits, ...tabDeclarations];
  }
}

function plannedNewTable(
  newTable: WideNewTable,
  placement: PlannedNewTablePlacement,
): PlannedNewTable {
  return { ...unplacedNewTable(newTable), placement };
}

// A new Table's entry named as the regen will name it, before it has a place.
export function unplacedNewTable({
  name,
  columns,
  conditionalFormats,
}: WideNewTable): Omit<PlannedNewTable, "placement"> {
  return {
    table: new SpreadsheetBaseSchema().titleToName(name),
    name,
    rules: conditionalFormats?.add ?? [],
    columns: columns.map(
      ({ header, emptyValueAllowed, formula, columnType }) => ({
        header,
        column: columnNameOfHeader(header),
        emptyValueAllowed: emptyValueAllowed ?? false,
        ...(formula === undefined ? {} : { formula }),
        ...(columnType === undefined ? {} : { columnType }),
      }),
    ),
  };
}

function onLiveTables<EN extends { table: string }>(
  entries: EN[],
  live: TraitReadings,
): EN[] {
  return entries.filter(({ table }) => live.tables[table] !== undefined);
}

interface SendLines<EN> {
  sendLine(entry: EN): SyncSendLine;
}

function effectlessSends<EN>(
  entries: EN[],
  lines: SendLines<EN>,
): SyncPlanItem[] {
  return entries.map((entry) => ({ ...lines.sendLine(entry), effects: [] }));
}

// A new tab's own line comes once, before its first Table's.
function newTableSends(newTables: PlannedNewTable[]): SyncPlanItem[] {
  const tabTitles = new Set<string>();
  return newTables.flatMap((newTable) => {
    const lines: SyncSendLine[] = [];
    const { placement } = newTable;
    if (placement.kind === "newTab" && !tabTitles.has(placement.title)) {
      tabTitles.add(placement.title);
      lines.push(newTableLines.tabLine(placement.title));
    }
    lines.push(newTableLines.sendLine(newTable));
    return lines.map((line) => ({ ...line, effects: [] }));
  });
}

interface ComparedSides {
  agreed: TraitValue | undefined;
  live: TraitValue | undefined;
  isColumnGone: boolean;
}

function hasRule(value: TraitValue | undefined): boolean {
  return typeof value === "object" && value.rule !== null;
}

function pickTraits<EN extends object, KY extends keyof EN & string>(
  entry: EN,
  traits: readonly KY[],
): { trait: string; value: TraitValue }[] {
  return traits.flatMap((trait) => {
    const value = entry[trait] as TraitValue | undefined;
    if (value === undefined) return [];
    return [{ trait, value }];
  });
}

// Order is never in the configs, so live alone decides it: sent unless live already has it.
function isOrderTrait(intended: IntendedTrait): boolean {
  return intended.kind !== "table" && intended.trait === "before";
}

export function isColumnDeletion(address: TraitAddress): boolean {
  return address.kind === "column" && address.trait === "deleted";
}

export function isTableDeletion(address: TraitAddress): boolean {
  return address.kind === "table" && address.trait === "deleted";
}

export function isTabDeletion(
  address: TraitAddress,
): address is TabTraitAddress & { trait: "deleted" } {
  return address.kind === "tab" && address.trait === "deleted";
}

function addressKey(address: TraitAddress): string {
  if (address.kind === "tab") {
    return JSON.stringify([address.kind, address.tab, address.trait]);
  }
  if (address.kind === "table") {
    return JSON.stringify([address.kind, address.table, address.trait]);
  }
  const { kind, table, column, trait } = address;
  return JSON.stringify([kind, table, column, trait]);
}

// Undefined when there is nothing to compare: a Table or tab a person deleted isn't a conflict.
function comparedSides(
  intended: IntendedTrait,
  previous: TraitReadings,
  live: TraitReadings,
): ComparedSides | undefined {
  const { trait } = intended;
  if (intended.kind === "tab") {
    const liveTab = live.tabs[intended.tab];
    if (liveTab === undefined) return undefined;
    return {
      agreed: previous.tabs[intended.tab]?.[trait] ?? liveTab[trait],
      live: liveTab[trait],
      isColumnGone: false,
    };
  }
  const liveTable = live.tables[intended.table];
  if (liveTable === undefined) return undefined;
  const previousTable = previous.tables[intended.table];
  if (intended.kind === "table") {
    return {
      agreed: previousTable?.traits[trait] ?? liveTable.traits[trait],
      live: liveTable.traits[trait],
      isColumnGone: false,
    };
  }
  const liveColumn = liveTable.columns[intended.column];
  const previousColumn = previousTable?.columns[intended.column];
  if (liveColumn === undefined && previousColumn === undefined) {
    return undefined;
  }
  return {
    agreed: previousColumn?.[trait] ?? liveColumn?.[trait],
    live: liveColumn?.[trait],
    isColumnGone: liveColumn === undefined,
  };
}

function stopOf(
  intended: IntendedTrait,
  { agreed, live, isColumnGone }: ComparedSides,
): StopContext | undefined {
  const base = { intended, agreed, live, isColumnGone };
  if (isColumnGone) return { ...base, ...stopWords.columnDeleted };
  if (isSameValue(live, agreed)) return undefined;
  if (isSameValue(intended.value, agreed)) {
    return { ...base, ...stopWords.personChanged };
  }
  if (isSameValue(intended.value, live)) return undefined;
  return { ...base, ...stopWords.bothChanged };
}

function sendsIntended(
  value: TraitValue,
  { agreed, live }: ComparedSides,
): boolean {
  return !isSameValue(value, agreed) && isSameValue(live, agreed);
}

function conflictOf(stop: StopContext): SyncConflict {
  const { intended, agreed, live, label, cause, isColumnGone } = stop;
  const conflict: SyncConflict = {
    kind: "conflict",
    subject: subjectOf(intended),
    label,
    cause,
    yours: valueText(intended.value),
    agreed: valueText(agreed),
    live: valueText(live),
  };
  if (isColumnGone) conflict.isColumnGone = true;
  return conflict;
}

export function subjectOf(address: TraitAddress): SyncSubject {
  if (address.kind === "tab") return { tab: address.tab, path: address.trait };
  if (address.kind === "table") {
    return { table: address.table, path: address.trait };
  }
  const { table, column, trait } = address;
  return { table, column, path: trait };
}

export function deletedTableOf(subject: SyncSubject): string | undefined {
  if (!("table" in subject) || subject.column !== undefined) return undefined;
  if (subject.path !== "deleted") return undefined;
  return subject.table;
}

export function deletedTabOf(subject: SyncSubject): string | undefined {
  if (!("tab" in subject) || subject.path !== "deleted") return undefined;
  return subject.tab;
}

// Both comparison and display, so a column type's key order can't read as a change.
export function valueText(value: TraitValue | undefined): string {
  if (value === undefined) return absentText;
  if (typeof value !== "object") return String(value);
  const type = value.type ?? "AUTOMATIC";
  if (value.rule === null) return type;
  return `${type} ${value.rule.condition} [${value.rule.values.join(", ")}]`;
}

export function isSameValue(
  first: TraitValue | undefined,
  second: TraitValue | undefined,
): boolean {
  return valueText(first) === valueText(second);
}

interface ConfigTraitsProps {
  live: TraitReadings;
  sent: IntendedTrait[];
  heldBack: HeldBackValue[];
  newColumns: PlannedNewColumn[];
  newTables: PlannedNewTable[];
  tabsByTable: TabsByTable;
}

function configTraitsOf({
  live,
  sent,
  heldBack,
  newColumns,
  newTables,
  tabsByTable,
}: ConfigTraitsProps): TraitReadings {
  const createdColumns: (PlannedNewTableColumn & { table: string })[] = [
    ...newColumns,
    ...newTables.flatMap(({ table, columns }) =>
      columns.map((column) => ({ table, ...column })),
    ),
  ];
  const created = createdColumns.flatMap((createdColumn) => {
    const { table, column, header, emptyValueAllowed, formula, columnType } =
      createdColumn;
    const traits: { trait: string; value: TraitValue | undefined }[] = [
      { trait: "header", value: header },
      { trait: "emptyValueAllowed", value: emptyValueAllowed },
      { trait: "formula", value: formula },
      { trait: "columnType", value: columnType },
    ];
    return traits.flatMap(({ trait, value }) => {
      if (value === undefined) return [];
      return [
        { address: { kind: "column" as const, table, column, trait }, value },
      ];
    });
  });
  const written = [
    ...sent.map((intended) => ({ address: intended, value: intended.value })),
    ...heldBack.map(({ intended, agreed }) => ({
      address: intended,
      value: agreed,
    })),
    ...created,
  ];
  return written.reduce((readings, { address, value }) => {
    if (isTabDeletion(address)) {
      return withoutTab(readings, address.tab, tabsByTable);
    }
    return withTrait(readings, address, value);
  }, live);
}

// A deleted tab takes its Tables with it.
function withoutTab(
  readings: TraitReadings,
  tab: string,
  tabsByTable: TabsByTable,
): TraitReadings {
  const { [tab]: _deleted, ...tabs } = readings.tabs;
  const tables = Object.fromEntries(
    Object.entries(readings.tables).filter(
      ([table]) => tabsByTable[table] !== tab,
    ),
  );
  return { tables, tabs };
}

function withTrait(
  readings: TraitReadings,
  address: TraitAddress,
  value: TraitValue,
): TraitReadings {
  const { trait } = address;
  if (address.kind === "tab") {
    const { tab } = address;
    return {
      ...readings,
      tabs: {
        ...readings.tabs,
        [tab]: { ...readings.tabs[tab], [trait]: value },
      },
    };
  }
  const { table } = address;
  if (isTableDeletion(address)) {
    const { [table]: _deleted, ...tables } = readings.tables;
    return { ...readings, tables };
  }
  const tableReading = readings.tables[table] ?? { traits: {}, columns: {} };
  if (address.kind === "table") {
    return withTable(readings, table, {
      ...tableReading,
      traits: { ...tableReading.traits, [trait]: value },
    });
  }
  const { column } = address;
  if (isColumnDeletion(address)) {
    const { [column]: _deleted, ...columns } = tableReading.columns;
    return withTable(readings, table, { ...tableReading, columns });
  }
  return withTable(readings, table, {
    ...tableReading,
    columns: {
      ...tableReading.columns,
      [column]: { ...tableReading.columns[column], [trait]: value },
    },
  });
}

function withTable(
  readings: TraitReadings,
  table: string,
  tableReading: TableReading,
): TraitReadings {
  return { ...readings, tables: { ...readings.tables, [table]: tableReading } };
}
