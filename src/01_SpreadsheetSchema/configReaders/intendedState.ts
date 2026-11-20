import type { ModelledConditionType } from "../../00_Source/RawSource/ConditionalFormat";
import type { ColumnTypeBaseline } from "../columnBaseline";
import type { IdNameMapRegistered } from "./configRegister";

type MapTables = IdNameMapRegistered["tables"];
type MapTabs = IdNameMapRegistered["tabs"];

// The map keys Tables by ID; intended state keys them as endpoint code names them.
type MapTablesByName = {
  [ID in keyof MapTables as MapTables[ID]["name"]]: MapTables[ID];
};

type MappedTableName = keyof MapTablesByName & string;

type MappedColumnName<TN extends MappedTableName> =
  MapTablesByName[TN] extends {
    columns: infer CS;
  }
    ? keyof CS & string
    : never;

// The regen writes each tab's name: camelCasing its title at the type level costs ~340 instantiations a title.
type MappedTabName = {
  [GID in keyof MapTabs]: MapTabs[GID]["name"];
}[keyof MapTabs];

// Deleting an entry says everything about it, so a trait beside the deletion can't be meant.
type Deletable<TR extends object> =
  | (TR & { deleted?: never })
  | ({ deleted: true } & { [K in keyof TR]?: never });

interface ColumnTraits {
  header?: string;
  // A dropdown carries its validation rule.
  columnType?: ColumnTypeBaseline;
  // Brace form: `${columnName}` for its own Table, `${tableName[columnName]}` for another.
  formula?: string;
  emptyValueAllowed?: boolean;
}

interface HeadedColumn extends ColumnTraits {
  header: string;
}

// Its columns, in order, sit together immediately before `before`, or at the Table's end without one; no two `newColumns` groups, nor a `moves` group, may name one `before`.
interface NewColumnGroup<CN extends string> {
  columns: HeadedColumn[];
  before?: CN;
}

// Its columns, in order, sit together immediately before `before`; no group's `before` may be a column any group moves, so every group can hold at once.
interface ColumnGroupMove<CN extends string> {
  columns: CN[];
  before: CN;
}

// Colours as hex text an agent writes, where the Source's `ConditionalFormat` holds RGB channels.
interface ConditionalFormatIntended {
  backgroundColor?: string;
  foregroundColor?: string;
  bold?: boolean;
  italic?: boolean;
  strikethrough?: boolean;
}

interface ConditionIntended {
  type: ModelledConditionType;
  // Each value in brace form, as a formula column's is.
  values?: string[];
}

// Spans the body rows of the columns named, so a rule never needs a column letter.
interface ConditionalFormatRuleIntended<CN extends string> {
  columns: CN[] | "all";
  condition: ConditionIntended;
  format: ConditionalFormatIntended;
}

type RuleHandle = string;

interface ConditionalFormatAlterations<CN extends string> {
  add?: ConditionalFormatRuleIntended<CN>[];
  remove?: RuleHandle[];
  replace?: Record<RuleHandle, ConditionalFormatRuleIntended<CN>>;
}

// Any string, while an editor still offers the mapped names it is unioned with.
type UnmappedName = string & Record<never, never>;

// The name `columnNameOfHeader` gives a `newColumns` header, which no type can compute; the entry refuser checks it before any read.
type NewColumnName = UnmappedName;

interface TableTraits<CN extends string, CE extends object> {
  // A Sheets Table's name is its codebase name, so a rename renames both.
  name?: string;
  letApiAccess?: boolean;
  columns?: { [C in CN]?: CE };
  newColumns?: NewColumnGroup<CN>[];
  // Applied in list order.
  moves?: ColumnGroupMove<CN>[];
  conditionalFormats?: ConditionalFormatAlterations<CN | NewColumnName>;
}

type TableAlteration<TN extends MappedTableName> = Deletable<
  TableTraits<MappedColumnName<TN>, Deletable<ColumnTraits>>
>;

// Its columns are named by header, since none is in the map until the run that adds it.
interface NewTable {
  name: string;
  columns: HeadedColumn[];
  conditionalFormats?: Pick<ConditionalFormatAlterations<string>, "add">;
}

interface TabTraits<TB extends string> {
  title?: string;
  color?: string;
  before?: TB;
  frozenRows?: number;
  frozenColumns?: number;
  hidden?: boolean;
}

interface TabAlteration extends TabTraits<MappedTabName> {
  // Placed right of the tab's rightmost Table, with `gap` empty columns between.
  newTables?: (NewTable & { gap?: 0 | 1 })[];
}

// A `newTabs` entry's title, which isn't in the map until the run that creates it.
type NewTabTitle = UnmappedName;

// Its `before` is sent after every addSheet of the run, so it may name another `newTabs` entry by title.
interface NewTab extends TabTraits<MappedTabName | NewTabTitle> {
  title: string;
  tables: NewTable[];
}

// Each entry is applied once, then removed.
export interface PendingAlterations {
  tables?: { [TN in MappedTableName]?: TableAlteration<TN> };
  tabs?: { [TB in MappedTabName]?: Deletable<TabAlteration> };
  newTabs?: NewTab[];
}

// Identity at runtime: it exists so `tsc` checks `alterations.ts` against the ID-name map before any send.
export function makePendingAlterations(
  pendingAlterations: PendingAlterations,
): PendingAlterations {
  return pendingAlterations;
}
