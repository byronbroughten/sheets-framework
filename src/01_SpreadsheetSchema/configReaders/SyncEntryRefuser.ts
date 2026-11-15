import { hasApiAccess } from "./apiAccessChanges";
import { BraceFormula } from "./BraceFormula";
import type { IdNameMapGeneric } from "./idNameMapTypes";
import { sheetTraits, tabTraits } from "./SyncPlanner";
import {
  entriesOf,
  type WideAlterations,
  type WideNewTable,
  type WideRule,
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
  table: new Set<string>([
    ...sheetTraits.table,
    "letApiAccess",
    "newColumns",
    "conditionalFormats.remove",
    "conditionalFormats.replace",
  ]),
  column: new Set<string>([...sheetTraits.column, "emptyValueAllowed"]),
  // A new column's own traits: a formula or column type would need a write after the insert.
  newColumn: new Set<string>(["header", "before", "emptyValueAllowed"]),
  tab: new Set<string>([...tabTraits, "newTables"]),
  // A new tab's other traits would need sheet-property updates after its addSheet.
  newTab: new Set<string>(["title", "tables"]),
  // A new Table's rules would need adds after its creation.
  newTable: new Set<string>(["name", "columns", "gap"]),
  newTableColumn: new Set<string>(["header", "emptyValueAllowed"]),
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
    const columnNamesByTable = new Map(
      mapTables.map(({ name, columns }) => [name, Object.keys(columns ?? {})]),
    );
    return entryRefusals({
      entry: "tables",
      mapNames: mapTables.map(({ name }) => name),
      entries: entriesOf(this.alterations.tables),
      traitRefusals: (table, entry) => {
        if (!apiAccessNames.has(table)) {
          return noApiAccessRefusals(table, entry);
        }
        return [
          ...tableTraitRefusals(table, entry),
          ...braceRefusals({ table, entry, columnNamesByTable }),
        ];
      },
    });
  }
  private get _tabRefusals(): SyncRefusal[] {
    return entryRefusals({
      entry: "tabs",
      mapNames: Object.values(this.idNameMap.tabs).map(({ name }) => name),
      entries: entriesOf(this.alterations.tabs),
      traitRefusals: (tab, entry) => [
        ...refusedPaths(tab, traitPathsAt("tab", entry, "")),
        ...newTablesRefusals(tab, entry.newTables ?? [], "newTables"),
      ],
    });
  }
  private get _newTabRefusals(): SyncRefusal[] {
    return (this.alterations.newTabs ?? []).flatMap((newTab, index) => {
      const at = `newTabs[${index}]`;
      const tab = newTab.title ?? "newTabs";
      // addSheet is sent only with a Table to create on it.
      const tablesRefusals =
        newTab.tables.length === 0
          ? [notSendableYet(tab, `${at}.tables`)]
          : newTablesRefusals(tab, newTab.tables, `${at}.tables`);
      return [
        ...refusedPaths(tab, traitPathsAt("newTab", newTab, `${at}.`)),
        ...tablesRefusals,
      ];
    });
  }
}

function refusedPaths(table: string, paths: TraitPath[]): SyncRefusal[] {
  return paths
    .filter(({ level, trait }) => !sendableTraits[level].has(trait))
    .map(({ path }) => notSendableYet(table, path));
}

// The working view doesn't list a Table its flush adds, so a second on the same tab would be placed over the first.
function newTablesRefusals(
  tab: string,
  newTables: WideNewTable[],
  tablesAt: string,
): SyncRefusal[] {
  const [first, ...others] = newTables;
  if (first === undefined) return [];
  return [
    ...refusedPaths(tab, newTableTraitPaths(first, `${tablesAt}[0]`)),
    ...others.map((_, index) =>
      notSendableYet(tab, `${tablesAt}[${index + 1}]`),
    ),
  ];
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

interface BraceRefusalsProps {
  table: string;
  entry: WideTableEntry;
  columnNamesByTable: Map<string, string[]>;
}

function braceRefusals({
  table,
  entry,
  columnNamesByTable,
}: BraceRefusalsProps): SyncRefusal[] {
  return entriesOf(entry.columns).flatMap(([column, { formula }]) => {
    if (formula === undefined) return [];
    const path = `${column}.formula`;
    const malformed = BraceFormula.malformedBraces({ formula }).map(
      (brace): SyncRefusal => ({ kind: "malformedBrace", table, path, brace }),
    );
    const unknown = BraceFormula.columnsNamed({
      tableKey: table,
      formula,
    }).flatMap(({ brace, tableKey, columnName }): SyncRefusal[] => {
      const columnNames = columnNamesByTable.get(tableKey);
      if (columnNames?.includes(columnName)) return [];
      return [
        {
          kind: "unknownBrace",
          table,
          path,
          brace,
          braceTable: tableKey,
          columns: columnNames,
        },
      ];
    });
    return [...malformed, ...unknown];
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
  return [
    ...refusedPaths(table, [
      ...traitPaths(entry),
      ...newColumnTraitPaths(entry),
    ]),
    ...ruleTextFormatPaths(entry).map((path) => notSendableYet(table, path)),
  ];
}

// The Source's ConditionalFormat holds only the two colours, so a text style has nowhere to go yet.
const ruleTextFormatTraits = ["bold", "italic", "strikethrough"] as const;

function ruleTextFormatPaths({ conditionalFormats }: WideTableEntry): string[] {
  const rules: [string, WideRule][] = [
    ...(conditionalFormats?.add ?? []).map(
      (rule, index): [string, WideRule] => [
        `conditionalFormats.add[${index}]`,
        rule,
      ],
    ),
    ...entriesOf(conditionalFormats?.replace).map(
      ([handle, rule]): [string, WideRule] => [
        `conditionalFormats.replace["${handle}"]`,
        rule,
      ],
    ),
  ];
  return rules.flatMap(([rulePath, { format }]) =>
    ruleTextFormatTraits.flatMap((trait) => {
      if (format[trait] === undefined) return [];
      return [`${rulePath}.format.${trait}`];
    }),
  );
}

function newColumnTraitPaths({ newColumns }: WideTableEntry): TraitPath[] {
  return (newColumns ?? []).flatMap((newColumn, index) =>
    traitPathsAt("newColumn", newColumn, `newColumns[${index}].`),
  );
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
