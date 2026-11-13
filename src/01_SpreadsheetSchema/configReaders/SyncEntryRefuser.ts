import { hasApiAccess } from "./apiAccessChanges";
import type { IdNameMapGeneric } from "./idNameMapTypes";
import { sheetTraits } from "./SyncPlanner";
import {
  entriesOf,
  type WideAlterations,
  type WideColumnEntry,
  type WideTableEntry,
} from "./SyncPlanner/wideAlterations";
import type { SyncRefusal } from "./syncRunResult";

export interface SyncEntryRefuserProps {
  alterations: WideAlterations;
  // The previous one: what the last sync recorded, since nothing is read yet.
  idNameMap: IdNameMapGeneric;
}

type NotSendableYet = Extract<SyncRefusal, { kind: "notSendableYet" }>;

// Only these traits can be sent so far; each later trait ticket takes its trait off this refusal.
export const sendableTraits = {
  table: new Set<string>([...sheetTraits.table, "letApiAccess"]),
  column: new Set<string>([...sheetTraits.column, "emptyValueAllowed"]),
} as const;

interface TraitPath {
  level: keyof typeof sendableTraits;
  trait: string;
  // The trait as a refusal names it: `column.trait` for a column's.
  path: string;
}

// Pure: refuses what can't be planned from intended state and the previous ID-name map alone, so before any read.
export class SyncEntryRefuser {
  private alterations: WideAlterations;
  private idNameMap: IdNameMapGeneric;
  constructor(props: SyncEntryRefuserProps) {
    this.alterations = props.alterations;
    this.idNameMap = props.idNameMap;
  }
  static init(props: SyncEntryRefuserProps): SyncEntryRefuser {
    return new SyncEntryRefuser(props);
  }
  get refusals(): SyncRefusal[] {
    return withoutRepeats([
      ...this._tableRefusals,
      ...this._tabRefusals,
      ...this._newTabRefusals,
    ]);
  }
  // Checked against the map, not the Table configs: those hold only Tables with API access.
  private get _tableRefusals(): SyncRefusal[] {
    const mapTables = Object.values(this.idNameMap.tables);
    // No type-level check that an entry names a Table with API access, for type-check cost; reconsider if the typing budget or the ID-name map's typing is revisited.
    const apiAccessNames = new Set(
      mapTables.flatMap((mapTable) => {
        if (!hasApiAccess(mapTable)) return [];
        return [mapTable.name];
      }),
    );
    return entryRefusals({
      entry: "tables",
      mapNames: mapTables.map(({ name }) => name),
      entries: entriesOf(this.alterations.tables),
      traitRefusals: (table, entry) => {
        if (apiAccessNames.has(table)) return tableTraitRefusals(table, entry);
        return noApiAccessRefusals(table, entry);
      },
    });
  }
  private get _tabRefusals(): SyncRefusal[] {
    return entryRefusals({
      entry: "tabs",
      mapNames: Object.values(this.idNameMap.tabs).map(({ name }) => name),
      entries: entriesOf(this.alterations.tabs),
      traitRefusals: (tab, entry) =>
        definedKeys(entry).map((trait) => notSendableYet(tab, trait)),
    });
  }
  private get _newTabRefusals(): SyncRefusal[] {
    return (this.alterations.newTabs ?? []).map(({ title }) =>
      notSendableYet(title ?? "newTabs", "newTabs"),
    );
  }
}

interface EntryRefusalsProps<EN extends object> {
  entry: "tables" | "tabs";
  mapNames: string[];
  entries: [string, EN][];
  traitRefusals: (name: string, entry: EN) => SyncRefusal[];
}

function entryRefusals<EN extends object>({
  entry,
  mapNames,
  entries,
  traitRefusals,
}: EntryRefusalsProps<EN>): SyncRefusal[] {
  const names = new Set(mapNames);
  return entries.flatMap(([name, traits]) => {
    if (!names.has(name)) return [noSuchTable(entry, name)];
    return traitRefusals(name, traits);
  });
}

function noSuchTable(entry: "tables" | "tabs", name: string): SyncRefusal {
  return { kind: "noSuchTable", entry, name };
}

function notSendableYet(table: string, path: string): NotSendableYet {
  return { kind: "notSendableYet", table, path };
}

function tableTraitRefusals(
  table: string,
  entry: WideTableEntry,
): SyncRefusal[] {
  return traitPaths(entry)
    .filter(({ level, trait }) => !sendableTraits[level].has(trait))
    .map(({ path }) => notSendableYet(table, path));
}

// A grant takes its own run, since a column entry needs the column IDs the grant writes.
function noApiAccessRefusals(
  table: string,
  entry: WideTableEntry,
): SyncRefusal[] {
  const kind = entry.letApiAccess === true ? "besideGrant" : "noApiAccess";
  return traitPaths(entry)
    .filter(({ path }) => path !== "letApiAccess")
    .map(({ path }) => ({ kind, table, path }));
}

function traitPaths({ columns, ...traits }: WideTableEntry): TraitPath[] {
  return [
    ...definedKeys(traits).map((trait): TraitPath => ({
      level: "table",
      trait,
      path: trait,
    })),
    ...entriesOf(columns).flatMap(([column, columnEntry]) =>
      columnTraitPaths(column, columnEntry),
    ),
  ];
}

function columnTraitPaths(column: string, entry: WideColumnEntry): TraitPath[] {
  return definedKeys(entry).map((trait) => ({
    level: "column",
    trait,
    path: `${column}.${trait}`,
  }));
}

function definedKeys(entry: object): string[] {
  return Object.entries(entry).flatMap(([key, value]) => {
    if (value === undefined) return [];
    return [key];
  });
}

// Alterations alone still repeat a refusal: two untitled `newTabs` are one.
function withoutRepeats(refusals: SyncRefusal[]): SyncRefusal[] {
  const seen = new Set<string>();
  return refusals.filter((refusal) => {
    const key = JSON.stringify(refusal);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}
