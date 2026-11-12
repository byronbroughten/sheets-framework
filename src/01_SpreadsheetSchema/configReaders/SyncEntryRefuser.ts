import type { IdNameMapGeneric } from "./idNameMapTypes";
import type { SyncClaim } from "./SyncPlanner";
import {
  entriesOf,
  type WideAlterations,
  type WideClaims,
  type WideColumnEntry,
  type WideTableEntry,
} from "./SyncPlanner/wideAlterations";
import type { SyncRefusal } from "./syncRunResult";

export interface SyncEntryRefuserProps {
  alterations: WideAlterations;
  claims: SyncClaim[];
  // The previous one: what the last sync recorded, since nothing is read yet.
  idNameMap: IdNameMapGeneric;
}

type NotSendableYet = Extract<SyncRefusal, { kind: "notSendableYet" }>;

// Only a column's emptyValueAllowed can be sent so far; each later trait ticket takes its trait off this refusal.
export const sendableColumnTraits = new Set<string>(["emptyValueAllowed"]);

// Pure: refuses what can't be planned from intended state and the previous ID-name map alone, so before any read.
export class SyncEntryRefuser {
  private alterations: WideAlterations;
  private claims: SyncClaim[];
  private idNameMap: IdNameMapGeneric;
  constructor(props: SyncEntryRefuserProps) {
    this.alterations = props.alterations;
    this.claims = props.claims;
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
      ...this._claimRefusals,
    ]);
  }
  // Checked against the map, not the Table configs: those hold only Tables with API access.
  private get _tableRefusals(): SyncRefusal[] {
    return entryRefusals({
      entry: "tables",
      mapNames: Object.values(this.idNameMap.tables).map(({ name }) => name),
      entries: entriesOf(this.alterations.tables),
      traitRefusals: tableTraitRefusals,
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
  private get _claimRefusals(): SyncRefusal[] {
    return this.claims.flatMap(({ traits }) =>
      entriesOf<WideClaims[string]>(traits).flatMap(([table, entry]) =>
        entriesOf(entry?.columns).flatMap(([column, claimed]) =>
          Object.keys(claimed).map((trait) =>
            notSendableYet(table, `${column}.${trait}`),
          ),
        ),
      ),
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
  const { columns, ...traits } = entry;
  return [
    ...definedKeys(traits).map((trait) => notSendableYet(table, trait)),
    ...entriesOf(columns).flatMap(([column, columnEntry]) =>
      columnTraitRefusals(table, column, columnEntry),
    ),
  ];
}

function columnTraitRefusals(
  table: string,
  column: string,
  entry: WideColumnEntry,
): SyncRefusal[] {
  return definedKeys(entry)
    .filter((trait) => !sendableColumnTraits.has(trait))
    .map((trait) => notSendableYet(table, `${column}.${trait}`));
}

function definedKeys(entry: object): string[] {
  return Object.entries(entry).flatMap(([key, value]) => {
    if (value === undefined) return [];
    return [key];
  });
}

// Two declarers of one trait, or a claim beside an alteration of it, are one refusal: the entry to take out is the same.
function withoutRepeats(refusals: SyncRefusal[]): SyncRefusal[] {
  const seen = new Set<string>();
  return refusals.filter((refusal) => {
    const key = JSON.stringify(refusal);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}
