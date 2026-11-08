import type { StrictOmit } from "@byronbroughten/utils/obj";

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

interface ColumnAlteration<CN extends string> extends ColumnTraits {
  before?: CN;
}

interface NewColumn<CN extends string> extends ColumnTraits {
  header: string;
  before?: CN;
}

interface NewTableColumn extends ColumnTraits {
  header: string;
}

// Colours as hex text an agent writes, where the Source's `ConditionalFormat` holds RGB channels.
interface ConditionalFormatIntended {
  backgroundColor?: string;
  foregroundColor?: string;
  bold?: boolean;
  italic?: boolean;
  strikethrough?: boolean;
}

// Spans the body rows of the columns named, so a rule never needs a column letter.
interface ConditionalFormatRuleIntended<CN extends string> {
  columns: CN[] | "all";
  // Each value in brace form, as a formula column's is.
  condition: { type: string; values?: string[] };
  format: ConditionalFormatIntended;
}

type RuleHandle = string;

interface ConditionalFormatAlterations<CN extends string> {
  add?: ConditionalFormatRuleIntended<CN>[];
  remove?: RuleHandle[];
  replace?: Record<RuleHandle, ConditionalFormatRuleIntended<CN>>;
}

interface TableTraits<CN extends string, CE extends object> {
  // A Sheets Table's name is its codebase name, so a rename renames both.
  name?: string;
  letApiAccess?: boolean;
  columns?: { [C in CN]?: CE };
  newColumns?: NewColumn<CN>[];
  conditionalFormats?: ConditionalFormatAlterations<CN>;
}

type TableAlteration<TN extends MappedTableName> = Deletable<
  TableTraits<
    MappedColumnName<TN>,
    Deletable<ColumnAlteration<MappedColumnName<TN>>>
  >
>;

// Its columns are named by header, since none is in the map until the run that adds it.
interface NewTable {
  name: string;
  columns: NewTableColumn[];
  conditionalFormats?: Pick<ConditionalFormatAlterations<string>, "add">;
}

interface TabTraits {
  title?: string;
  color?: string;
  before?: MappedTabName;
  frozenRows?: number;
  frozenColumns?: number;
  hidden?: boolean;
}

interface TabAlteration extends TabTraits {
  // Placed right of the tab's rightmost Table, with `gap` empty columns between.
  newTables?: (NewTable & { gap?: 0 | 1 })[];
}

interface NewTab extends TabTraits {
  title: string;
  tables: NewTable[];
}

// Each entry is applied once, then removed; a claim is compared every run instead.
export interface PendingAlterations {
  tables?: { [TN in MappedTableName]?: TableAlteration<TN> };
  tabs?: { [TB in MappedTabName]?: Deletable<TabAlteration> };
  newTabs?: NewTab[];
}

// Formula text only, for now: no other trait has a comparison to stop a run on yet.
type ClaimedColumnTraits = Pick<ColumnTraits, "formula">;

type ClaimedTableTraits<TN extends MappedTableName> = StrictOmit<
  TableTraits<MappedColumnName<TN>, ClaimedColumnTraits>,
  "newColumns" | "name" | "letApiAccess" | "conditionalFormats"
>;

export type ClaimedTraits = {
  [TN in MappedTableName]?: ClaimedTableTraits<TN>;
};

// Identity at runtime: it exists so `tsc` checks `alterations.ts` against the ID-name map before any send.
export function makePendingAlterations(
  pendingAlterations: PendingAlterations,
): PendingAlterations {
  return pendingAlterations;
}
