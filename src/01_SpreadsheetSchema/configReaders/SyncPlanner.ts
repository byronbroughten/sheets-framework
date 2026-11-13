import { Str } from "@byronbroughten/utils/str";

import type { ColumnTypeBaseline } from "../columnBaseline";
import type { PendingAlterations } from "./intendedState";
import { DeletionRefuser } from "./SyncPlanner/DeletionRefuser";
import { entriesOf, type WideAlterations } from "./SyncPlanner/wideAlterations";
import type {
  ForcedTrait,
  SyncConflict,
  SyncHeldBackTrait,
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

export interface SyncPlannerProps {
  alterations: PendingAlterations;
  tabsByTable: Record<string, string | undefined>;
  // Only ever this run's --force arguments, so a yes can't outlive its run.
  forcedTraits: ForcedTrait[];
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
  column: string;
  header: string;
  before?: string;
  emptyValueAllowed: boolean;
}

export interface SyncPlannerResult extends Pick<
  SyncRunResult,
  "stops" | "plan"
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

// The sent traits Raw's working view doesn't carry, so the run refetches after sending one; a tab's land in the working view.
export const sheetTraits = {
  table: new Set<string>(["name"]),
  column: new Set<string>(["header", "formula", "deleted"]),
} as const;

// A deletion's one value, `true`, says nothing a plan line could show.
const deletionWhat = "delete the column";

// The subject path of every `newColumns` entry's plan line; landing one drops the whole list from alterations.ts.
export const newColumnsPath = "newColumns";

// Where a new column goes, as its plan line and its created line say it.
export function newColumnPosition({ before }: PlannedNewColumn): string {
  if (before === undefined) return "at the Table's end";
  return `before ${before}`;
}

// A new column's plan line, and its didntLand line when the refetch can't find it.
export function newColumnSendLine(newColumn: PlannedNewColumn): SyncSendLine {
  return {
    subject: { table: newColumn.table, path: newColumnsPath },
    what: `create "${newColumn.header}" ${newColumnPosition(newColumn)}`,
  };
}

export function isSheetTrait(intended: IntendedTrait): boolean {
  if (intended.kind === "tab") return false;
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
  private tabsByTable: Record<string, string | undefined>;
  private forcedTraits: ForcedTrait[];
  constructor(props: SyncPlannerProps) {
    this.alterations = props.alterations;
    this.tabsByTable = props.tabsByTable;
    this.forcedTraits = props.forcedTraits;
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
        ({ header, before, emptyValueAllowed }): PlannedNewColumn[] => {
          if (header === undefined) return [];
          return [
            {
              table,
              column: Str.sentenceToCamelCase(header),
              header,
              ...(before === undefined ? {} : { before }),
              emptyValueAllowed: emptyValueAllowed ?? false,
            },
          ];
        },
      ),
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
      function send(what: string): void {
        sends.push({ subject: subjectOf(intended), what, effects: [] });
        sentTraits.push(intended);
      }
      if (isOrderTrait(intended)) {
        if (!isSameValue(intended.value, sides.live)) {
          send(`set to ${valueText(intended.value)}`);
        }
        return;
      }
      const stop = stopOf(intended, sides);
      if (stop === undefined) {
        if (!sendsIntended(intended.value, sides)) return;
        if (isColumnDeletion(intended)) {
          send(deletionWhat);
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
    const newColumns = this.newColumns.filter(
      ({ table }) => live.tables[table] !== undefined,
    );
    for (const newColumn of newColumns) {
      sends.push({ ...newColumnSendLine(newColumn), effects: [] });
    }
    const isStopped = stops.length + sheetStops.length > 0;
    const sentIfUnstopped = isStopped ? [] : sentTraits;
    const newColumnsIfUnstopped = isStopped ? [] : newColumns;
    const forcedNotHeld = new Set(
      forced.map((trait) => addressKey({ kind: "column", ...trait })),
    );
    const heldBackTraits = heldBack.filter(
      ({ intended }) => isStopped || !forcedNotHeld.has(addressKey(intended)),
    );
    return {
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
      }),
      sentTraits: sentIfUnstopped,
      newColumns: newColumnsIfUnstopped,
    };
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
        ...pickTraits(entry, ["name", "letApiAccess"]).map((trait) => ({
          kind: "table" as const,
          table,
          ...trait,
        })),
        ...entriesOf(entry.columns).flatMap(([column, columnEntry]) =>
          pickTraits(columnEntry, [
            "header",
            "columnType",
            "formula",
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
        pickTraits(entry, tabTraits).map((trait) => ({
          kind: "tab",
          tab,
          ...trait,
        })),
    );
    return [...tableTraits, ...tabDeclarations];
  }
}

interface ComparedSides {
  agreed: TraitValue | undefined;
  live: TraitValue | undefined;
  isColumnGone: boolean;
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
  return intended.kind === "tab" && intended.trait === "before";
}

export function isColumnDeletion(address: TraitAddress): boolean {
  return address.kind === "column" && address.trait === "deleted";
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
}

function configTraitsOf({
  live,
  sent,
  heldBack,
  newColumns,
}: ConfigTraitsProps): TraitReadings {
  const created = newColumns.flatMap(
    ({ table, column, header, emptyValueAllowed }) =>
      [
        { trait: "header", value: header },
        { trait: "emptyValueAllowed", value: emptyValueAllowed },
      ].map(({ trait, value }) => ({
        address: { kind: "column" as const, table, column, trait },
        value,
      })),
  );
  const written = [
    ...sent.map((intended) => ({ address: intended, value: intended.value })),
    ...heldBack.map(({ intended, agreed }) => ({
      address: intended,
      value: agreed,
    })),
    ...created,
  ];
  return written.reduce(
    (readings, { address, value }) => withTrait(readings, address, value),
    live,
  );
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
