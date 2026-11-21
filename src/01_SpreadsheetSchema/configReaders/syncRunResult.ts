// What one `sync` run did or would do, as `SyncRunReport` prints it. Facts, apart from the few phrases a producer words: a stop's label and cause, a send's what, a sent summary.

export type SyncRunResult = SyncDryRunResult | SyncSendResult;

export interface SyncDryRunResult extends SyncRunCommon {
  mode: "dry run";
  outcome: SyncDryRunOutcome;
  held?: SyncHeld;
}

export interface SyncSendResult extends SyncRunCommon {
  mode: "send";
  outcome: SyncSendOutcome;
  landed: SyncLandedItem[];
  notSent: SyncSendLine[];
  didntLand: SyncSendLine[];
  // From the file the run leaves, since an entry the planner skips lists in no section yet stays in it; false when no entry landed or was met, as on a Row-ID-only send or one that writes no files.
  hasEmptiedAlterations: boolean;
}

// What a run's caller says about it, before anything is read.
export type SyncRunFrame = Pick<
  SyncRunCommon,
  "target" | "command" | "generatedDir" | "alterationsPath"
>;

export interface SyncRunCommon {
  target: string;
  // The dry run's command, e.g. `npm run app:sync`; Next names it.
  command: string;
  generatedDir: string;
  alterationsPath: string;
  checks: SyncCheck[];
  refusals: SyncRefusal[];
  stops: SyncStop[];
  plan: SyncPlanItem[];
  // Alterations the configs and the live sheet both already hold: never sent, and removed from alterations.ts by an unstopped send as if landed.
  alreadyMet: SyncSendLine[];
  // The stops this run's --force turned into sends; the entered-data gate holds each.
  forcedTraits: ForcedTrait[];
  notices: SyncNotice[];
  configs: SyncConfigs;
}

export type SyncDryRunOutcome =
  { kind: "stopped" } | { kind: "refused" } | { kind: "ready" };

export type SyncSendOutcome =
  | { kind: "stopped" }
  | { kind: "refused" }
  | SyncSentOutcome
  | SyncFlushStoppedOutcome;

// A send stopped at a flush the entered-data gate held, Google refused or, for the fill's own send, failed.
export type SyncFlushStoppedOutcome =
  SyncStoppedBeforeSendingOutcome | SyncStoppedPartwayOutcome;

export interface SyncSentOutcome {
  kind: "sent";
  summary: string;
}

export type SyncStoppedBeforeSendingOutcome = {
  kind: "stoppedBeforeSending";
} & SyncStopCause;

export type SyncStoppedPartwayOutcome = {
  kind: "stoppedPartway";
} & SyncLandedFlushes &
  SyncStopCause;

export interface SyncLandedFlushes {
  // At least 1 on a partway stop: a send stopped at its first flush is stoppedBeforeSending, since nothing in it landed.
  landedFlushes: number;
  // The dry run's flush count, from its held file; absent when the send had none to read.
  flushCount?: number;
}

export type SyncStopCause = SyncRefusedCause | SyncFillFailedCause;

// The gate held the flush, or Google refused its batch whole, in the reworded text that names the Tables.
export type SyncRefusedCause =
  { heldBy: SyncHeldBy } | { googleRefusal: string };

// The Fill Row IDs send after the regen threw, by its error's message; the run still writes its files.
export interface SyncFillFailedCause {
  fillFailure: string;
}

// "uncovered": the send had no --confirm; "mismatched": the sheet or plan changed since the dry run.
export type SyncHeldBy = "uncovered" | "mismatched";

// Names a Table, a column or a tab by codebase name, as alterations.ts keys it.
export type SyncSubject = SyncTableSubject | SyncTabSubject;

export interface SyncTableSubject {
  table: string;
  column?: string;
  rewrittenFrom?: SyncRewrittenFrom;
  path?: string;
}

// Set when sync rewrote an entry's key: the name the entry was written with, and the column it still names.
export interface SyncRewrittenFrom {
  wrote: string;
  columnId: string;
}

// "blank": a deleted Table's outside references read blank cells; "ref": a deleted tab's show #REF!.
export type SyncOutsideReferenceEffect = "blank" | "ref";

export interface SyncTabSubject {
  tab: string;
  gid?: number;
  // A tab's own trait, as a Table subject's `path` holds one of a Table's.
  path?: string;
}

export interface SyncSendLine {
  subject: SyncSubject;
  what: string;
}

export type SyncLandedItem =
  ({ kind: "send" } & SyncSendLine) | SyncCreatedColumn | SyncCreatedTable;

// A `newTabs` or `newTables` Table as the refetch found it, by the codebase name its name gives.
export interface SyncCreatedTable {
  kind: "createdTable";
  table: string;
  tab: string;
  tableId: string;
  columns: SyncCreatedTableColumn[];
}

export interface SyncCreatedTableColumn {
  header: string;
  column: string;
  columnId: string;
}

export interface SyncCreatedColumn extends SyncCreatedTableColumn {
  kind: "createdColumn";
  table: string;
}

export interface SyncOutsideReference {
  cell: string;
  formula: string;
}

export interface SyncOutsideReferencesCheck {
  kind: "outsideReferences";
  subject: SyncSubject;
  effect: SyncOutsideReferenceEffect;
  references: SyncOutsideReference[];
}

export type SyncCheck =
  | SyncOutsideReferencesCheck
  | {
      kind: "followedSwap";
      subject: SyncTableSubject & { rewrittenFrom: SyncRewrittenFrom };
      header: string;
    }
  | {
      kind: "handDrawnRule";
      table: string;
      tab: string;
      range: string;
      formula: string;
      handle: string;
      cause: string;
    }
  | { kind: "runNotice"; text: string };

export type SyncEntry = "tables" | "tabs" | "newTabs";

export type SyncRefusal =
  // An entry on a Table without API access that this run doesn't grant it.
  | { kind: "noApiAccess"; table: string; path: string }
  // Any other entry on a Table this run grants API access.
  | { kind: "besideGrant"; table: string; path: string }
  // A `tables` or `tabs` entry naming one the previous ID-name map doesn't hold.
  | { kind: "noSuchTable"; entry: Exclude<SyncEntry, "newTabs">; name: string }
  // `name` is the Table's, the tab's or the new tab's title, as `entry` says; `path` is the trait's, with its column in front for a column's.
  | {
      kind: "notSendableYet";
      entry: SyncEntry;
      name: string;
      path: string;
    }
  // A `newTabs` entry's `before` naming neither a tab the previous ID-name map holds nor a `newTabs` title; `table` is the entry's title.
  | { kind: "noSuchBeforeTab"; table: string; path: string; before: string }
  // A brace naming a Table the previous ID-name map lacks.
  | {
      kind: "unknownBraceTable";
      table: string;
      path: string;
      brace: string;
      braceTable: string;
    }
  // A brace naming a column of braceTable the previous ID-name map lacks, or for a rule's brace one the Table's `newColumns` create; `columns` is braceTable's.
  | {
      kind: "unknownBrace";
      table: string;
      path: string;
      brace: string;
      braceTable: string;
      columns: string[];
    }
  // A `${…}` in a formula that reads as no brace, such as `@` before a Table key.
  | { kind: "malformedBrace"; table: string; path: string; brace: string }
  // A rule's column neither in the map nor created by the Table's `newColumns`; `columns` lists both.
  | {
      kind: "unknownColumn";
      table: string;
      path: string;
      column: string;
      columns: string[];
    }
  // A `newColumns` or new Table header giving the column name an earlier one's does, at `firstPath`: the refetch can't tell the two apart.
  | {
      kind: "repeatedNewColumn";
      table: string;
      path: string;
      firstPath: string;
      column: string;
    }
  // A `conditionalFormats.replace` handle the Table's `remove` also names, at `removedAt`.
  | {
      kind: "removedAndReplaced";
      table: string;
      path: string;
      removedAt: string;
    }
  | SyncClashingPlacements
  // A rule naming a column by the name a rename in this run gives it, while alterations key it by `oldName` (ADR 0035).
  | {
      kind: "renamedColumn";
      table: string;
      path: string;
      column: string;
      oldName: string;
    }
  | {
      kind: "touchesDeletion";
      deleted: SyncSubject;
      deletedAt: string;
      touchedBy: string;
      // Set when touchedBy deletes this Table or one of its columns, which the tab's deletion already takes.
      takenTable?: string;
    };

// Two `moves` or `newColumns` paths that can't both hold: one `before` twice, one column in two groups, or a `before` some group moves.
export interface SyncClashingPlacements {
  kind: "clashingPlacements";
  clash:
    | "sameBefore"
    | "repeatedColumn"
    | "movedBefore"
    | "sameNewColumnsBefore"
    | "newColumnsBeforeMoveTarget"
    | "newColumnsBeforeMovedColumn";
  table: string;
  path: string;
  otherPath: string;
  column: string;
}

export interface ForcedTrait {
  table: string;
  column: string;
  trait: string;
}

export type SyncStop = SyncConflict | SyncSheetStop;

// A stop found reading the sheet, outside the three-way comparison.
export type SyncSheetStop =
  | SyncHeadRowsStop
  | SyncHeaderTakenStop
  | SyncBeforeFirstColumnStop
  | SyncNoSuchRuleStop
  | SyncUnmodelableRuleStop
  | SyncNewTableStop;

export type SyncNewTableStop =
  SyncTableNameTakenStop | SyncTabTitleTakenStop | SyncTabTitleRepeatedStop;

// Two or more `newTabs` entries giving one codebase name: `paths` are theirs, as `newTabs[0]`.
export interface SyncTabTitleRepeatedStop {
  kind: "tabTitleRepeated";
  title: string;
  paths: string[];
}

// A new Table whose name gives the codebase name a live Table's does: `existing` and `tableId` are that Table's.
export interface SyncTableNameTakenStop {
  kind: "tableNameTaken";
  name: string;
  existing: string;
  tableId: string;
}

// A `newTabs` title giving the codebase name a live tab's does: `existing` and `gid` are that tab's.
export interface SyncTabTitleTakenStop {
  kind: "tabTitleTaken";
  title: string;
  existing: string;
  gid: number;
}

export interface SyncConflict {
  kind: "conflict";
  subject: SyncSubject;
  label: string;
  cause: string;
  yours: string;
  agreed: string;
  live: string;
  // Set when a person deleted the column: force has no column to send to.
  isColumnGone?: true;
}

// A grant whose Table can't take head rows: its header sits off the header zone, or its head rows would sit on another Table.
export type SyncHeadRowsStop = SyncHeaderOffZoneStop | SyncHeadRowsOverlapStop;

interface SyncHeadRowsStopBase {
  kind: "headRows";
  table: string;
  tab: string;
}

export interface SyncHeaderOffZoneStop extends SyncHeadRowsStopBase {
  headerRow: number;
}

// `neighbour` is the live name of the Table the head rows would sit on.
export interface SyncHeadRowsOverlapStop extends SyncHeadRowsStopBase {
  neighbour: string;
}

// A `newColumns` entry whose header names a column the Table already has: `column` and `columnId` are that column's.
export interface SyncHeaderTakenStop {
  kind: "headerTaken";
  table: string;
  header: string;
  column: string;
  columnId: string;
}

// A `newColumns` entry whose `before` names its Table's first column by the live order: an insert there lands outside the Table.
export interface SyncBeforeFirstColumnStop {
  kind: "beforeFirstColumn";
  table: string;
  path: string;
  header: string;
  before: string;
}

// The one phrase for it, in the stop's block and in the throw guarding the insert.
export function beforeFirstColumnText({
  table,
  header,
  before,
}: Omit<SyncBeforeFirstColumnStop, "kind" | "path">): string {
  return `Table ${table} can't take "${header}" before ${before}, its first column: a column inserted there lands outside the Table.`;
}

// A `conditionalFormats.remove` or `replace` handle no rule on the Table has on this run's read; `handles` are the Table's rules' now.
export interface SyncNoSuchRuleStop {
  kind: "noSuchRule";
  table: string;
  path: string;
  handles: string[];
}

// A `conditionalFormats.replace` handle naming a rule the configs record as unmodelable, which sync can only remove.
export interface SyncUnmodelableRuleStop {
  kind: "unmodelableRule";
  table: string;
  path: string;
  handle: string;
}

export interface SyncHeld {
  groups: SyncHeldGroup[];
  token: string;
  fullListPath: string;
}

export interface SyncHeldGroup {
  // Absent for a held item no Table or tab names, such as a find-and-replace.
  subject?: SyncSubject;
  what: string;
  // Entered-data cells this group holds; absent for a held item that isn't cells, such as a dropdown reset.
  cells?: number;
  lines: string[];
}

export interface SyncPlanItem extends SyncSendLine {
  effects: SyncPlanEffect[];
}

export type SyncPlanEffect =
  | { kind: "dropdownReset"; columns: string[]; setting: "ask" | "allow" }
  | { kind: "addedFirst"; outranks: string[] }
  | {
      kind: "rulesRewritten";
      table: string;
      count: number;
      after: "insert" | "move" | "delete";
    }
  | { kind: "ruleRespanned"; handle: string; fromRow: number; toRow: number }
  | { kind: "rulesRemoved"; handles: string[] }
  | { kind: "rulePartlyOver"; tab: string; range: string; formula: string }
  | {
      kind: "outsideReferences";
      count: number;
      effect: SyncOutsideReferenceEffect;
    }
  | { kind: "overwrites"; cells: number; blank: number };

export type SyncNotice =
  | { kind: "idFirst"; table: string }
  | SyncTableGoneNotice
  | SyncTabGoneNotice
  | SyncRuleHeldBackNotice
  | { kind: "text"; text: string };

// A rule add or replace, or a column change with its rule rewrites, on a Table whose columns `shiftedBy`'s column insert or delete shifts this run; it stays in alterations.ts and is sent next run.
export interface SyncHeldBackChange {
  subject: SyncTableSubject;
  shiftedBy: string;
}

// A dry run's word for a held-back rule or column change, which a send lists as not sent.
export interface SyncRuleHeldBackNotice extends SyncHeldBackChange {
  kind: "ruleHeldBack";
}

// One phrase for the not-sent line and the dry run's notice.
export function heldBackChangeWhat({
  subject,
  shiftedBy,
}: SyncHeldBackChange): string {
  return `held back: ${shiftedBy}'s column change shifts ${subject.table}'s columns this run, so it is sent next run`;
}

// A Table the previous configs hold that a person deleted from the sheet.
export interface SyncTableGoneNotice {
  kind: "tableGone";
  table: string;
}

// A tab the previous ID-name map holds that a person deleted from the sheet.
export interface SyncTabGoneNotice {
  kind: "tabGone";
  tab: string;
}

// The framework endpoint's name, in every plan, summary and notice line that cites it.
export const fillRowIdsName = "Fill Row IDs";

export interface SyncConfigs {
  heldBack: SyncHeldBackTrait[];
  renamed: SyncRename[];
  namesMoved: SyncMovedName[];
  rewrites: SyncRewrite[];
  typedTraitChanges: SyncTypedTraitChange[];
  notes: string[];
}

export interface SyncHeldBackTrait {
  subject: SyncSubject;
  value: string;
}

export interface SyncRename {
  table: string;
  from: string;
  to: string;
  headerFrom: string;
  headerTo: string;
}

export interface SyncMovedTableOrTabName {
  kind: "table" | "tab";
  name: string;
  wasId: string;
  nowId: string;
}

export interface SyncMovedColumnName {
  kind: "column";
  table: string;
  column: string;
  wasId: string;
  nowId: string;
}

export type SyncMovedName = SyncMovedTableOrTabName | SyncMovedColumnName;

export interface SyncRewrite {
  from: string;
  to: string;
  why?: string;
}

export interface SyncTypedTraitChange {
  table: string;
  column: string;
  trait: string;
  from: string;
  to: string;
}
