import type { ColumnTypeBaseline } from "../01_SpreadsheetSchema/columnBaseline";
import { BraceFormula } from "../01_SpreadsheetSchema/configReaders/BraceFormula";
import type {
  ClaimedTraits,
  PendingAlterations,
} from "../01_SpreadsheetSchema/configReaders/intendedState";
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

interface TraitAddress {
  table?: string;
  tab?: string;
  column?: string;
  trait: string;
}

interface IntendedTrait extends TraitAddress {
  value: TraitValue;
  claimedBy?: string;
}

interface Declaration extends IntendedTrait {
  declaredBy: string;
}

interface StopContext {
  intended: IntendedTrait;
  agreed: TraitValue | undefined;
  live: TraitValue | undefined;
  label: string;
  cause: string;
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
    this.alterations = props.alterations as WideAlterations;
    this.tabsByTable = props.tabsByTable;
    this.forcedTraits = props.forcedTraits;
  }
  static init(props: SyncPlannerProps): SyncPlanner {
    return new SyncPlanner(props);
  }
  // Takes only intended state, so the run can refuse before reading anything.
  get refusals(): SyncRefusal[] {
    return [...this._disagreeingDeclarers, ...this._touchedDeletions];
  }
  plan({ previous, live }: PlanProps): SyncPlannerResult {
    const stops: SyncConflict[] = [];
    const sends: SyncPlanItem[] = [];
    const forced: ForcedTrait[] = [];
    const sentTraits: IntendedTrait[] = [];
    const heldBack: { intended: IntendedTrait; agreed: TraitValue }[] = [];
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
    const forcedNotHeld = new Set(forced.map(forcedKey));
    const heldBackTraits = heldBack.filter(
      ({ intended }) =>
        isStopped || !forcedNotHeld.has(forcedKey(forcedFrom(intended))),
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
    return this.forcedTraits.find(
      ({ table, column, trait }) =>
        table === intended.table &&
        column === intended.column &&
        trait === intended.trait,
    );
  }
  private get _claimDeclarations(): Declaration[] {
    return this.claims.flatMap(({ declaredBy, traits }) =>
      Object.entries(traits as WideClaims).flatMap(([table, entry]) =>
        Object.entries(entry?.columns ?? {}).flatMap(([column, claimed]) => {
          if (claimed?.formula === undefined) return [];
          return [
            {
              table,
              column,
              trait: "formula",
              value: claimed.formula,
              claimedBy: declaredBy,
              declaredBy,
            },
          ];
        }),
      ),
    );
  }
  private get _alterationDeclarations(): Declaration[] {
    const declaredBy = alterationsDeclarer;
    const tableTraits = entriesOf(this.alterations.tables).flatMap(
      ([table, entry]) => [
        ...pickTraits(entry, ["name", "letApiAccess"]).map((trait) => ({
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
          ]).map((trait) => ({ table, column, ...trait, declaredBy })),
        ),
      ],
    );
    const tabTraits = entriesOf(this.alterations.tabs).flatMap(([tab, entry]) =>
      pickTraits(entry, [
        "title",
        "color",
        "frozenRows",
        "frozenColumns",
        "hidden",
      ]).map((trait) => ({ tab, ...trait, declaredBy })),
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
      const { declaredBy: _declaredBy, ...trait } = first;
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
        const table = first.table ?? first.tab ?? "";
        const path = [first.column, first.trait]
          .filter((part) => part !== undefined)
          .join(".");
        const declarers: [SyncDeclarer, SyncDeclarer] = [
          declarerOf(first),
          declarerOf(other),
        ];
        return [
          {
            kind: "disagreeingDeclarers" as const,
            table,
            path,
            declarers,
          },
        ];
      });
    });
  }
  private get _touchedDeletions(): SyncRefusal[] {
    const touches = [...this._alterationTouches, ...this._claimTouches];
    return this._deletions.flatMap((deletion) =>
      touches.flatMap(({ targets, touchedBy }) => {
        if (!targets.includes(deletion.target)) return [];
        return [
          {
            kind: "touchesDeletion" as const,
            deleted: deletion.subject,
            deletedAt: deletion.deletedAt,
            touchedBy,
          },
        ];
      }),
    );
  }
  private get _deletions(): Deletion[] {
    const tableDeletions = entriesOf(this.alterations.tables).flatMap(
      ([table, entry]) => [
        ...(entry.deleted === true
          ? [
              {
                target: tableTarget(table),
                subject: { table },
                deletedAt: `tables.${table}`,
              },
            ]
          : []),
        ...entriesOf(entry.columns).flatMap(([column, columnEntry]) => {
          if (columnEntry.deleted !== true) return [];
          return [
            {
              target: columnTarget(table, column),
              subject: { table, column },
              deletedAt: `tables.${table}.columns.${column}`,
            },
          ];
        }),
      ],
    );
    const tabDeletions = entriesOf(this.alterations.tabs).flatMap(
      ([tab, entry]) => {
        if (entry.deleted !== true) return [];
        return [
          {
            target: tabTarget(tab),
            subject: { tab },
            deletedAt: `tabs.${tab}`,
          },
        ];
      },
    );
    return [...tableDeletions, ...tabDeletions];
  }
  private get _claimTouches(): Touch[] {
    return this._claimDeclarations.flatMap(
      ({ table, column, value, declaredBy }) => {
        if (table === undefined || column === undefined) return [];
        const touchedBy = `claim (${declaredBy})`;
        return [
          {
            targets: this._columnTargets(table, column),
            touchedBy,
          },
          ...this._braceTouches(table, String(value), touchedBy),
        ];
      },
    );
  }
  private get _alterationTouches(): Touch[] {
    const tableTouches = entriesOf(this.alterations.tables).flatMap(
      ([table, entry]) => this._tableEntryTouches(table, entry),
    );
    const tabTouches = entriesOf(this.alterations.tabs).flatMap(
      ([tab, entry]) => {
        if (entry.deleted === true) return [];
        const at = `tabs.${tab}`;
        return [
          { targets: [tabTarget(tab)], touchedBy: at },
          ...beforeTabTouches(entry.before, `${at}.before`),
          ...this._newTablesTouches(entry.newTables ?? [], `${at}.newTables`),
        ];
      },
    );
    const newTabTouches = (this.alterations.newTabs ?? []).flatMap(
      (newTab, index) => {
        const at = `newTabs[${index}]`;
        return [
          ...beforeTabTouches(newTab.before, `${at}.before`),
          ...this._newTablesTouches(newTab.tables, `${at}.tables`),
        ];
      },
    );
    return [...tableTouches, ...tabTouches, ...newTabTouches];
  }
  private _tableEntryTouches(table: string, entry: WideTableEntry): Touch[] {
    if (entry.deleted === true) return [];
    const at = `tables.${table}`;
    return [
      { targets: this._tableTargets(table), touchedBy: at },
      ...entriesOf(entry.columns).flatMap(([column, columnEntry]) =>
        this._columnEntryTouches(table, column, columnEntry),
      ),
      ...(entry.newColumns ?? []).flatMap((newColumn, index) =>
        this._placedColumnTouches(
          table,
          newColumn,
          `${at}.newColumns[${index}]`,
        ),
      ),
      ...this._ruleTouches(table, entry.conditionalFormats, at),
    ];
  }
  private _columnEntryTouches(
    table: string,
    column: string,
    entry: WideColumnEntry,
  ): Touch[] {
    if (entry.deleted === true) return [];
    const at = `tables.${table}.columns.${column}`;
    return [
      { targets: this._columnTargets(table, column), touchedBy: at },
      ...this._placedColumnTouches(table, entry, at),
    ];
  }
  private _placedColumnTouches(
    table: string,
    entry: WideColumnEntry,
    at: string,
  ): Touch[] {
    const { before, formula } = entry;
    return [
      ...(before === undefined
        ? []
        : [
            {
              targets: this._columnTargets(table, before),
              touchedBy: `${at}.before`,
            },
          ]),
      ...(formula === undefined
        ? []
        : this._braceTouches(table, formula, `${at}.formula`)),
    ];
  }
  private _ruleTouches(
    table: string,
    formats: WideTableEntry["conditionalFormats"],
    tableAt: string,
  ): Touch[] {
    if (formats === undefined) return [];
    const at = `${tableAt}.conditionalFormats`;
    const added = (formats.add ?? []).map(
      (rule, index) => [`${at}.add[${index}]`, rule] as const,
    );
    const replaced = entriesOf(formats.replace).map(
      ([handle, rule]) => [`${at}.replace["${handle}"]`, rule] as const,
    );
    return [...added, ...replaced].flatMap(([ruleAt, rule]) => [
      ...(rule.columns === "all"
        ? []
        : rule.columns.map((column) => ({
            targets: this._columnTargets(table, column),
            touchedBy: `${ruleAt}.columns`,
          }))),
      ...(rule.condition.values ?? []).flatMap((value) =>
        this._braceTouches(table, value, `${ruleAt}.condition`),
      ),
    ]);
  }
  private _newTablesTouches(
    newTables: WideNewTable[],
    tablesAt: string,
  ): Touch[] {
    return newTables.flatMap((newTable, tableIndex) =>
      newTable.columns.flatMap(({ formula }, columnIndex) => {
        if (formula === undefined) return [];
        return this._braceTouches(
          newTable.name,
          formula,
          `${tablesAt}[${tableIndex}].columns[${columnIndex}].formula`,
        );
      }),
    );
  }
  private _braceTouches(
    tableKey: string,
    formula: string,
    touchedBy: string,
  ): Touch[] {
    return BraceFormula.columnsNamed({ tableKey, formula }).map(
      ({ tableKey: table, columnName }) => ({
        targets: this._columnTargets(table, columnName),
        touchedBy,
      }),
    );
  }
  private _columnTargets(table: string, column: string): string[] {
    return [columnTarget(table, column), ...this._tableTargets(table)];
  }
  private _tableTargets(table: string): string[] {
    const tab = this.tabsByTable[table];
    return [tableTarget(table), ...(tab === undefined ? [] : [tabTarget(tab)])];
  }
}

// `PendingAlterations` keyed by whatever the map holds, widened since the planner reads every entry alike.
interface WideRule {
  columns: string[] | "all";
  condition: { values?: string[] };
}

interface WideColumnEntry {
  deleted?: true;
  before?: string;
  formula?: string;
  header?: string;
  columnType?: ColumnTypeBaseline;
  emptyValueAllowed?: boolean;
}

interface WideTableEntry {
  deleted?: true;
  name?: string;
  letApiAccess?: boolean;
  columns?: Record<string, WideColumnEntry | undefined>;
  newColumns?: WideColumnEntry[];
  conditionalFormats?: {
    add?: WideRule[];
    replace?: Record<string, WideRule | undefined>;
  };
}

interface WideNewTable {
  name: string;
  columns: { formula?: string }[];
}

interface WideTabEntry {
  deleted?: true;
  before?: string;
  title?: string;
  color?: string;
  frozenRows?: number;
  frozenColumns?: number;
  hidden?: boolean;
  newTables?: WideNewTable[];
}

interface WideAlterations {
  tables?: Record<string, WideTableEntry | undefined>;
  tabs?: Record<string, WideTabEntry | undefined>;
  newTabs?: (WideTabEntry & { tables: WideNewTable[] })[];
}

type WideClaims = Record<
  string,
  { columns?: Record<string, { formula?: string } | undefined> } | undefined
>;

interface Deletion {
  target: string;
  subject: SyncSubject;
  deletedAt: string;
}

interface Touch {
  targets: string[];
  touchedBy: string;
}

interface ComparedSides {
  agreed: TraitValue | undefined;
  live: TraitValue | undefined;
  isColumnGone: boolean;
}

function entriesOf<VL>(
  record: Record<string, VL | undefined> | undefined,
): [string, VL][] {
  return Object.entries(record ?? {}).flatMap(([key, value]) => {
    if (value === undefined) return [];
    return [[key, value] as [string, VL]];
  });
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

function beforeTabTouches(
  before: string | undefined,
  touchedBy: string,
): Touch[] {
  if (before === undefined) return [];
  return [{ targets: [tabTarget(before)], touchedBy }];
}

function columnTarget(table: string, column: string): string {
  return `column:${table}/${column}`;
}

function tableTarget(table: string): string {
  return `table:${table}`;
}

function tabTarget(tab: string): string {
  return `tab:${tab}`;
}

function addressKey({ table, tab, column, trait }: TraitAddress): string {
  return JSON.stringify([table, tab, column, trait]);
}

function declarerOf({ declaredBy, value }: Declaration): SyncDeclarer {
  return { declaredBy, value: valueText(value) };
}

// Undefined when there is nothing to compare: a Table or tab a person deleted isn't a conflict.
function comparedSides(
  intended: IntendedTrait,
  previous: TraitReadings,
  live: TraitReadings,
): ComparedSides | undefined {
  const { tab, table, column, trait } = intended;
  if (tab !== undefined) {
    const liveTab = live.tabs[tab];
    if (liveTab === undefined) return undefined;
    return {
      agreed: previous.tabs[tab]?.[trait] ?? liveTab[trait],
      live: liveTab[trait],
      isColumnGone: false,
    };
  }
  if (table === undefined) return undefined;
  const liveTable = live.tables[table];
  if (liveTable === undefined) return undefined;
  const previousTable = previous.tables[table];
  if (column === undefined) {
    return {
      agreed: previousTable?.traits[trait] ?? liveTable.traits[trait],
      live: liveTable.traits[trait],
      isColumnGone: false,
    };
  }
  const liveColumn = liveTable.columns[column];
  const previousColumn = previousTable?.columns[column];
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
  const base = { intended, agreed, live };
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
  const { intended, agreed, live, label, cause } = stop;
  return {
    kind: "conflict",
    subject: subjectOf(intended),
    label,
    cause,
    yours: valueText(intended.value),
    agreed: valueText(agreed),
    live: valueText(live),
    ...(intended.claimedBy === undefined
      ? {}
      : { claimedBy: intended.claimedBy }),
  };
}

function subjectOf({ tab, table, column, trait }: IntendedTrait): SyncSubject {
  if (tab !== undefined) return { tab, path: trait };
  return {
    table: table ?? "",
    ...(column === undefined ? {} : { column }),
    path: trait,
  };
}

function forcedFrom({ table, column, trait }: IntendedTrait): ForcedTrait {
  return { table: table ?? "", column: column ?? "", trait };
}

function forcedKey({ table, column, trait }: ForcedTrait): string {
  return JSON.stringify([table, column, trait]);
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
  heldBack: { intended: IntendedTrait; agreed: TraitValue }[];
}

function configTraitsOf({
  live,
  sent,
  heldBack,
}: ConfigTraitsProps): TraitReadings {
  const configTraits: TraitReadings = copyOfReadings(live);
  sent.forEach((intended) => {
    writeTrait(configTraits, intended, intended.value);
  });
  heldBack.forEach(({ intended, agreed }) => {
    writeTrait(configTraits, intended, agreed);
  });
  return configTraits;
}

function copyOfReadings({ tables, tabs }: TraitReadings): TraitReadings {
  return {
    tables: Object.fromEntries(
      entriesOf(tables).map(([table, { traits, columns }]) => [
        table,
        {
          traits: { ...traits },
          columns: Object.fromEntries(
            entriesOf(columns).map(([column, reading]) => [
              column,
              { ...reading },
            ]),
          ),
        },
      ]),
    ),
    tabs: Object.fromEntries(
      entriesOf(tabs).map(([tab, reading]) => [tab, { ...reading }]),
    ),
  };
}

function writeTrait(
  readings: TraitReadings,
  { tab, table, column, trait }: TraitAddress,
  value: TraitValue,
): void {
  if (tab !== undefined) {
    readings.tabs[tab] = { ...readings.tabs[tab], [trait]: value };
    return;
  }
  if (table === undefined) return;
  const tableReading = readings.tables[table] ?? { traits: {}, columns: {} };
  readings.tables[table] = tableReading;
  if (column === undefined) {
    tableReading.traits[trait] = value;
    return;
  }
  tableReading.columns[column] = {
    ...tableReading.columns[column],
    [trait]: value,
  };
}
