import type { IdNameMapGeneric } from "./idNameMapTypes";
import type { SyncClaim } from "./SyncPlanner";
import {
  entriesOf,
  type WideAlterations,
  type WideClaims,
  type WideColumnEntry,
  type WideTabEntry,
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
const sendableColumnTraits = new Set<string>(["emptyValueAllowed"]);

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
    return [
      ...this._tableRefusals,
      ...this._tabRefusals,
      ...this._newTabRefusals,
      ...this._claimRefusals,
    ];
  }
  // Checked against the map, not the Table configs: those hold only Tables with API access.
  private get _tableRefusals(): SyncRefusal[] {
    const names = new Set(
      Object.values(this.idNameMap.tables).map(({ name }) => name),
    );
    return entriesOf(this.alterations.tables).flatMap(([name, entry]) => {
      if (!names.has(name)) return [noSuchTable("tables", name)];
      return tableTraitRefusals(name, entry);
    });
  }
  private get _tabRefusals(): SyncRefusal[] {
    const names = new Set(
      Object.values(this.idNameMap.tabs).map(({ name }) => name),
    );
    return entriesOf(this.alterations.tabs).flatMap(([name, entry]) => {
      if (!names.has(name)) return [noSuchTable("tabs", name)];
      return tabTraitRefusals(name, entry);
    });
  }
  private get _newTabRefusals(): SyncRefusal[] {
    return (this.alterations.newTabs ?? []).map(({ title }) =>
      notSendableYet(title ?? "newTabs", "newTabs"),
    );
  }
  private get _claimRefusals(): SyncRefusal[] {
    const refusals = this.claims.flatMap(({ traits }) =>
      entriesOf<WideClaims[string]>(traits).flatMap(([table, entry]) =>
        entriesOf(entry?.columns).flatMap(([column, claimed]) =>
          Object.keys(claimed).map((trait) =>
            notSendableYet(table, `${column}.${trait}`),
          ),
        ),
      ),
    );
    return withoutRepeats(refusals);
  }
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

function tabTraitRefusals(tab: string, entry: WideTabEntry): SyncRefusal[] {
  return definedKeys(entry).map((trait) => notSendableYet(tab, trait));
}

function definedKeys(entry: object): string[] {
  return Object.entries(entry).flatMap(([key, value]) => {
    if (value === undefined) return [];
    return [key];
  });
}

// Two declarers claiming one trait are one refusal: the entry to take out is the same.
function withoutRepeats(refusals: NotSendableYet[]): NotSendableYet[] {
  const seen = new Set<string>();
  return refusals.filter(({ table, path }) => {
    const key = `${table}\n${path}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}
