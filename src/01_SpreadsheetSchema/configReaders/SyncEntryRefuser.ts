import { hasApiAccess } from "./apiAccessChanges";
import type { IdNameMapGeneric } from "./idNameMapTypes";
import { BraceRefuser, type NewTableAt } from "./SyncEntryRefuser/BraceRefuser";
import { NewColumnNames } from "./SyncEntryRefuser/NewColumnNames";
import { PlacementsRefuser } from "./SyncEntryRefuser/PlacementsRefuser";
import { RenamedColumnNames } from "./SyncEntryRefuser/RenamedColumnNames";
import { RuleRefuser } from "./SyncEntryRefuser/RuleRefuser";
import { sheetTraits, tabTraits } from "./SyncPlanner";
import {
  entriesOf,
  newColumnGroupColumnPath,
  newColumnPath,
  rulesOf,
  type WideAlterations,
  type WideNewTable,
  type WideTableEntry,
} from "./SyncPlanner/wideAlterations";
import type { SyncEntry, SyncRefusal } from "./syncRunResult";

export interface SyncEntryRefuserProps {
  alterations: WideAlterations;
  // The previous one: what the last sync recorded, since nothing is read yet.
  idNameMap: IdNameMapGeneric;
}

type NotSendableYet = Extract<SyncRefusal, { kind: "notSendableYet" }>;
type TraitEntry = Pick<NotSendableYet, "entry" | "name">;
type TabTraitEntry = TraitEntry & { entry: Exclude<SyncEntry, "tables"> };
type MapEntry = Extract<SyncRefusal, { kind: "noSuchTable" }>["entry"];
type NoSuchBeforeTab = Extract<SyncRefusal, { kind: "noSuchBeforeTab" }>;

// Only these traits can be sent so far; each later trait ticket takes its trait off this refusal.
export const sendableTraits = {
  table: new Set<string>([
    ...sheetTraits.table,
    "letApiAccess",
    "newColumns",
    "moves",
    "conditionalFormats.remove",
    "conditionalFormats.replace",
    "conditionalFormats.add",
  ]),
  column: new Set<string>([...sheetTraits.column, "emptyValueAllowed"]),
  // A new column's own traits: a formula or column type would need a write after the insert.
  newColumn: new Set<string>(["header", "emptyValueAllowed"]),
  newColumnGroup: new Set<string>(["columns", "before"]),
  tab: new Set<string>([...tabTraits, ...sheetTraits.tab, "newTables"]),
  newTab: new Set<string>([...tabTraits, "tables"]),
  newTable: new Set<string>([
    "name",
    "columns",
    "gap",
    "conditionalFormats.add",
  ]),
  newTableColumn: new Set<string>([
    "header",
    "emptyValueAllowed",
    "formula",
    "columnType",
  ]),
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
    const columnNamesByTable = this._columnNamesByTable;
    const renamedColumnsByTable = RenamedColumnNames.byTable(this.alterations);
    const braceRefuser = this._braceRefuser;
    return entryRefusals({
      entry: "tables",
      mapNames: mapTables.map(({ name }) => name),
      entries: entriesOf(this.alterations.tables),
      traitRefusals: (table, entry) => {
        if (!apiAccessNames.has(table)) {
          return noApiAccessRefusals(table, entry);
        }
        const newColumns = NewColumnNames.ofTableEntry(entry);
        return [
          ...tableTraitRefusals(table, entry),
          ...newColumns.repeatRefusals(table),
          ...new PlacementsRefuser(table, entry).refusals,
          ...braceRefuser.tableEntryRefusals(table, entry),
          ...RuleRefuser.ofTableEntry({
            table,
            entry,
            newColumns,
            columnNamesByTable,
            renamedColumnsByTable,
          }).refusals,
        ];
      },
    });
  }
  private get _braceRefuser(): BraceRefuser {
    return new BraceRefuser({ columnNamesByTable: this._columnNamesByTable });
  }
  private get _columnNamesByTable(): Map<string, string[]> {
    return new Map(
      Object.values(this.idNameMap.tables).map(({ name, columns }) => [
        name,
        Object.keys(columns ?? {}),
      ]),
    );
  }
  private get _tabRefusals(): SyncRefusal[] {
    return entryRefusals({
      entry: "tabs",
      mapNames: Object.values(this.idNameMap.tabs).map(({ name }) => name),
      entries: entriesOf(this.alterations.tabs),
      traitRefusals: (tab, entry) => {
        const tabEntry: TabTraitEntry = { entry: "tabs", name: tab };
        return [
          ...refusedPaths(tabEntry, traitPathsAt("tab", entry, "")),
          ...this._newTablesRefusals({
            tabEntry,
            newTables: entry.newTables ?? [],
            tablesAt: "newTables",
          }),
        ];
      },
    });
  }
  private get _newTabRefusals(): SyncRefusal[] {
    const newTabs = this.alterations.newTabs ?? [];
    // A new tab's traits go out after every addSheet, so a `before` the sheet can't resolve then fails with a write already landed.
    const beforeTabs = new Set([
      ...Object.values(this.idNameMap.tabs).map(({ name }) => name),
      ...newTabs.map(({ title }) => title),
    ]);
    return newTabs.flatMap((newTab, index) => {
      const at = `newTabs[${index}]`;
      const tab = newTab.title;
      const tabEntry: TabTraitEntry = { entry: "newTabs", name: tab };
      // addSheet is sent only with a Table to create on it.
      const tablesRefusals =
        newTab.tables.length === 0
          ? [notSendableYet({ ...tabEntry, path: `${at}.tables` })]
          : this._newTablesRefusals({
              tabEntry,
              newTables: newTab.tables,
              tablesAt: `${at}.tables`,
            });
      const { before } = newTab;
      const beforeRefusals =
        before === undefined || beforeTabs.has(before)
          ? []
          : [noSuchBeforeTab({ table: tab, path: `${at}.before`, before })];
      return [
        ...refusedPaths(tabEntry, traitPathsAt("newTab", newTab, `${at}.`)),
        ...beforeRefusals,
        ...tablesRefusals,
      ];
    });
  }
  private _newTablesRefusals({
    tabEntry,
    newTables,
    tablesAt,
  }: NewTablesRefusalsProps): SyncRefusal[] {
    return newTables.flatMap((newTable, index) =>
      this._newTableRefusals({
        tabEntry,
        newTable,
        at: `${tablesAt}[${index}]`,
      }),
    );
  }
  private _newTableRefusals({
    tabEntry,
    newTable,
    at,
  }: NewTableUnder): SyncRefusal[] {
    const tab = tabEntry.name;
    const newColumns = NewColumnNames.ofNewTable({ newTable, at });
    return [
      ...refusedPaths(tabEntry, newTableTraitPaths(newTable, at)),
      ...newColumns.repeatRefusals(tab),
      ...this._braceRefuser.newTableRefusals({ tab, newTable, at }),
      ...RuleRefuser.ofNewTable({
        tab,
        newTable,
        at,
        newColumns,
        columnNamesByTable: this._columnNamesByTable,
        renamedColumnsByTable: RenamedColumnNames.byTable(this.alterations),
      }).refusals,
      ...ruleTextFormatPaths({
        conditionalFormats: newTable.conditionalFormats,
      }).map((path) => notSendableYet({ ...tabEntry, path: `${at}.${path}` })),
    ];
  }
}

interface NewTableUnder extends Omit<NewTableAt, "tab"> {
  tabEntry: TabTraitEntry;
}

interface NewTablesRefusalsProps {
  tabEntry: TabTraitEntry;
  newTables: WideNewTable[];
  tablesAt: string;
}

function refusedPaths(
  traitEntry: TraitEntry,
  paths: TraitPath[],
): SyncRefusal[] {
  return paths
    .filter(({ level, trait }) => !sendableTraits[level].has(trait))
    .map(({ path }) => notSendableYet({ ...traitEntry, path }));
}

function newTableTraitPaths(
  { columns, conditionalFormats, ...traits }: WideNewTable,
  at: string,
): TraitPath[] {
  return [
    ...traitPathsAt("newTable", { ...traits, columns }, `${at}.`),
    ...definedKeys(conditionalFormats ?? {}).map((operation) => ({
      level: "newTable" as const,
      trait: `conditionalFormats.${operation}`,
      path: `${at}.conditionalFormats.${operation}`,
    })),
    ...columns.flatMap((column, columnIndex) =>
      traitPathsAt("newTableColumn", column, `${at}.columns[${columnIndex}].`),
    ),
  ];
}

interface EntryRefusalsProps<EN extends object> {
  entry: MapEntry;
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

function noSuchTable(entry: MapEntry, name: string): SyncRefusal {
  return { kind: "noSuchTable", entry, name };
}

function notSendableYet(props: Omit<NotSendableYet, "kind">): NotSendableYet {
  return { kind: "notSendableYet", ...props };
}

function noSuchBeforeTab(
  props: Omit<NoSuchBeforeTab, "kind">,
): NoSuchBeforeTab {
  return { kind: "noSuchBeforeTab", ...props };
}

function tableTraitRefusals(
  table: string,
  entry: WideTableEntry,
): SyncRefusal[] {
  const traitEntry: TraitEntry = { entry: "tables", name: table };
  return [
    ...refusedPaths(traitEntry, [
      ...traitPaths(entry),
      ...newColumnTraitPaths(entry),
    ]),
    ...ruleTextFormatPaths(entry).map((path) =>
      notSendableYet({ ...traitEntry, path }),
    ),
  ];
}

// The Source's ConditionalFormat holds only the two colours, so a text style has nowhere to go yet.
const ruleTextFormatTraits = ["bold", "italic", "strikethrough"] as const;

function ruleTextFormatPaths(entry: WideTableEntry): string[] {
  return rulesOf(entry).flatMap(({ path: rulePath, rule: { format } }) =>
    ruleTextFormatTraits.flatMap((trait) => {
      if (format[trait] === undefined) return [];
      return [`${rulePath}.format.${trait}`];
    }),
  );
}

function newColumnTraitPaths({ newColumns }: WideTableEntry): TraitPath[] {
  return (newColumns ?? []).flatMap((item, index) => {
    const at = `${newColumnPath(index)}.`;
    return [
      ...traitPathsAt("newColumnGroup", item, at),
      ...item.columns.flatMap((column, columnIndex) =>
        traitPathsAt(
          "newColumn",
          column,
          `${newColumnGroupColumnPath({ index, columnIndex })}.`,
        ),
      ),
    ];
  });
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

function traitPaths({
  columns,
  conditionalFormats,
  ...traits
}: WideTableEntry): TraitPath[] {
  return [
    ...definedKeys(traits).map(tableTraitPath),
    // Each rule operation is its own trait, since sync sends some before others.
    ...definedKeys(conditionalFormats ?? {}).map((operation) =>
      tableTraitPath(`conditionalFormats.${operation}`),
    ),
    ...entriesOf(columns).flatMap(([column, columnEntry]) =>
      traitPathsAt("column", columnEntry, `${column}.`),
    ),
  ];
}

function tableTraitPath(trait: string): TraitPath {
  return { level: "table", trait, path: trait };
}

// `prefix` is where the entry sits in its alterations, ending in a dot, or "" for a top-level entry.
function traitPathsAt(
  level: TraitPath["level"],
  entry: object,
  prefix: string,
): TraitPath[] {
  return definedKeys(entry).map((trait) => ({
    level,
    trait,
    path: `${prefix}${trait}`,
  }));
}

function definedKeys(entry: object): string[] {
  return Object.entries(entry).flatMap(([key, value]) => {
    if (value === undefined) return [];
    return [key];
  });
}

// Alterations alone can still repeat a refusal, so each is listed once.
function withoutRepeats(refusals: SyncRefusal[]): SyncRefusal[] {
  const seen = new Set<string>();
  return refusals.filter((refusal) => {
    const key = JSON.stringify(refusal);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}
