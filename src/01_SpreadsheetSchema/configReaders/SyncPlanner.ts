import type { ColumnTypeBaseline } from "../columnBaseline";
import type { ClaimedTraits, PendingAlterations } from "./intendedState";
import {
  type ClaimedFormula,
  DeletionRefuser,
} from "./SyncPlanner/DeletionRefuser";
import {
  entriesOf,
  type WideAlterations,
  type WideClaims,
} from "./SyncPlanner/wideAlterations";
import type {
  ForcedTrait,
  SyncConflict,
  SyncDeclarer,
  SyncHeldBackTrait,
  SyncPlanItem,
  SyncRefusal,
  SyncRunResult,
  SyncSubject,
} from "./syncRunResult";

export type TraitValue = string | number | boolean | ColumnTypeBaseline;
type TraitRecord = Record<string, TraitValue | undefined>;

interface TableReading {
  traits: TraitRecord;
  columns: Record<string, TraitRecord | undefined>;
}

// One side of the comparison: the previous configs or the live sheet, by codebase name.
export interface TraitReadings {
  tables: Record<string, TableReading | undefined>;
  tabs: Record<string, TraitRecord | undefined>;
}

export interface SyncClaim {
  declaredBy: string;
  traits: ClaimedTraits;
}

export interface SyncPlannerProps {
  claims: SyncClaim[];
  alterations: PendingAlterations;
  tabsByTable: Record<string, string | undefined>;
  // Only ever this run's --force arguments, so a yes can't outlive its run.
  forcedTraits: ForcedTrait[];
}

interface PlanProps {
  previous: TraitReadings;
  live: TraitReadings;
}

export interface SyncPlannerResult extends Pick<
  SyncRunResult,
  "stops" | "plan"
> {
  configs: { heldBack: SyncHeldBackTrait[] };
  // Handed to the flush classifier, which makes each one a held item.
  forcedTraits: ForcedTrait[];
  // What the regenerated configs record if the plan's sends land.
  configTraits: TraitReadings;
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
  claimedBy?: string;
}

type IntendedTrait = TraitAddress & IntendedValue;

type Declaration = IntendedTrait & { declaredBy: string };

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

const alterationsDeclarer = "alterations.ts";

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
  private claims: SyncClaim[];
  private alterations: WideAlterations;
  private tabsByTable: Record<string, string | undefined>;
  private forcedTraits: ForcedTrait[];
  constructor(props: SyncPlannerProps) {
    this.claims = props.claims;
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
      claimedFormulas: this._claimedFormulas,
    });
  }
  // Takes only intended state, so the run can refuse before reading anything.
  get refusals(): SyncRefusal[] {
    return [...this._disagreeingDeclarers, ...this._deletionRefuser.refusals];
  }
  plan({ previous, live }: PlanProps): SyncPlannerResult {
    const stops: SyncConflict[] = [];
    const sends: SyncPlanItem[] = [];
    const forced: ForcedTrait[] = [];
    const sentTraits: IntendedTrait[] = [];
    const heldBack: HeldBackValue[] = [];
    this._intendedTraits.forEach((intended) => {
      const sides = comparedSides(intended, previous, live);
      if (sides === undefined) return;
      const stop = stopOf(intended, sides);
      function send(what: string): void {
        sends.push({ subject: subjectOf(intended), what, effects: [] });
        sentTraits.push(intended);
      }
      if (stop === undefined) {
        if (sendsIntended(intended.value, sides)) {
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
    const isStopped = stops.length > 0;
    const forcedNotHeld = new Set(
      forced.map((trait) => addressKey({ kind: "column", ...trait })),
    );
    const heldBackTraits = heldBack.filter(
      ({ intended }) => isStopped || !forcedNotHeld.has(addressKey(intended)),
    );
    return {
      stops,
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
        sent: isStopped ? [] : sentTraits,
        heldBack: heldBackTraits,
      }),
    };
  }
  private _forcedTraitFor(intended: IntendedTrait): ForcedTrait | undefined {
    const key = addressKey(intended);
    return this.forcedTraits.find(
      (trait) => addressKey({ kind: "column", ...trait }) === key,
    );
  }
  private get _claimedFormulas(): ClaimedFormula[] {
    return this.claims.flatMap(({ declaredBy, traits }) =>
      entriesOf<WideClaims[string]>(traits).flatMap(([table, entry]) =>
        entriesOf(entry?.columns).flatMap(([column, claimed]) => {
          if (claimed.formula === undefined) return [];
          return [{ table, column, formula: claimed.formula, declaredBy }];
        }),
      ),
    );
  }
  private get _claimDeclarations(): Declaration[] {
    return this._claimedFormulas.map(
      ({ table, column, formula, declaredBy }) => ({
        kind: "column",
        table,
        column,
        trait: "formula",
        value: formula,
        claimedBy: declaredBy,
        declaredBy,
      }),
    );
  }
  private get _alterationDeclarations(): Declaration[] {
    const declaredBy = alterationsDeclarer;
    const tableTraits = entriesOf(this.alterations.tables).flatMap(
      ([table, entry]): Declaration[] => [
        ...pickTraits(entry, ["name", "letApiAccess"]).map((trait) => ({
          kind: "table" as const,
          table,
          ...trait,
          declaredBy,
        })),
        ...entriesOf(entry.columns).flatMap(([column, columnEntry]) =>
          pickTraits(columnEntry, [
            "header",
            "columnType",
            "formula",
            "emptyValueAllowed",
          ]).map((trait) => ({
            kind: "column" as const,
            table,
            column,
            ...trait,
            declaredBy,
          })),
        ),
      ],
    );
    const tabTraits = entriesOf(this.alterations.tabs).flatMap(
      ([tab, entry]): Declaration[] =>
        pickTraits(entry, [
          "title",
          "color",
          "frozenRows",
          "frozenColumns",
          "hidden",
        ]).map((trait) => ({ kind: "tab", tab, ...trait, declaredBy })),
    );
    return [...tableTraits, ...tabTraits];
  }
  private get _declarationsByTrait(): Map<string, Declaration[]> {
    return [...this._claimDeclarations, ...this._alterationDeclarations].reduce(
      (byTrait, declaration) => {
        const key = addressKey(declaration);
        byTrait.set(key, [...(byTrait.get(key) ?? []), declaration]);
        return byTrait;
      },
      new Map<string, Declaration[]>(),
    );
  }
  // A claim's declarer stays on the merged trait: removing the alteration wouldn't clear its conflict.
  private get _intendedTraits(): IntendedTrait[] {
    return [...this._declarationsByTrait.values()].flatMap((declarations) => {
      const [first] = declarations;
      if (first === undefined) return [];
      const claimedBy = declarations.find(
        ({ claimedBy: claimant }) => claimant !== undefined,
      )?.claimedBy;
      const trait = intendedOf(first);
      if (claimedBy === undefined) return [trait];
      return [{ ...trait, claimedBy }];
    });
  }
  private get _disagreeingDeclarers(): SyncRefusal[] {
    return [...this._declarationsByTrait.values()].flatMap((declarations) => {
      const [first, ...rest] = declarations;
      if (first === undefined) return [];
      return rest.flatMap((other) => {
        if (isSameValue(first.value, other.value)) return [];
        const declarers: [SyncDeclarer, SyncDeclarer] = [
          declarerOf(first),
          declarerOf(other),
        ];
        return [
          {
            kind: "disagreeingDeclarers" as const,
            ...refusalPlaceOf(first),
            declarers,
          },
        ];
      });
    });
  }
}

interface ComparedSides {
  agreed: TraitValue | undefined;
  live: TraitValue | undefined;
  isColumnGone: boolean;
}

function pickTraits<EN extends object, KY extends keyof EN & string>(
  entry: EN,
  traits: KY[],
): { trait: string; value: TraitValue }[] {
  return traits.flatMap((trait) => {
    const value = entry[trait] as TraitValue | undefined;
    if (value === undefined) return [];
    return [{ trait, value }];
  });
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

function intendedOf(declaration: Declaration): IntendedTrait {
  const { declaredBy: _declaredBy, ...trait } = declaration;
  return trait;
}

function declarerOf({ declaredBy, value }: Declaration): SyncDeclarer {
  return { declaredBy, value: valueText(value) };
}

function refusalPlaceOf(address: TraitAddress): {
  table: string;
  path: string;
} {
  if (address.kind === "tab") {
    return { table: address.tab, path: address.trait };
  }
  if (address.kind === "table") {
    return { table: address.table, path: address.trait };
  }
  return { table: address.table, path: `${address.column}.${address.trait}` };
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
  if (intended.claimedBy !== undefined) {
    conflict.claimedBy = intended.claimedBy;
  }
  if (isColumnGone) conflict.isColumnGone = true;
  return conflict;
}

function subjectOf(address: TraitAddress): SyncSubject {
  if (address.kind === "tab") return { tab: address.tab, path: address.trait };
  if (address.kind === "table") {
    return { table: address.table, path: address.trait };
  }
  const { table, column, trait } = address;
  return { table, column, path: trait };
}

// Both comparison and display, so a column type's key order can't read as a change.
function valueText(value: TraitValue | undefined): string {
  if (value === undefined) return absentText;
  if (typeof value !== "object") return String(value);
  const type = value.type ?? "AUTOMATIC";
  if (value.rule === null) return type;
  return `${type} ${value.rule.condition} [${value.rule.values.join(", ")}]`;
}

function isSameValue(
  first: TraitValue | undefined,
  second: TraitValue | undefined,
): boolean {
  return valueText(first) === valueText(second);
}

interface ConfigTraitsProps {
  live: TraitReadings;
  sent: IntendedTrait[];
  heldBack: HeldBackValue[];
}

function configTraitsOf({
  live,
  sent,
  heldBack,
}: ConfigTraitsProps): TraitReadings {
  const written = [
    ...sent.map((intended) => ({ address: intended, value: intended.value })),
    ...heldBack.map(({ intended, agreed }) => ({
      address: intended,
      value: agreed,
    })),
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
