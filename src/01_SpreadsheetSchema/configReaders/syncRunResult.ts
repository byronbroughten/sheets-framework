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
  | SyncHeldStoppedOutcome;

// A send the entered-data gate refused at a held flush.
export type SyncHeldStoppedOutcome =
  SyncStoppedBeforeSendingOutcome | SyncStoppedPartwayOutcome;

export interface SyncSentOutcome {
  kind: "sent";
  summary: string;
}

export interface SyncStoppedBeforeSendingOutcome {
  kind: "stoppedBeforeSending";
  heldBy: SyncHeldBy;
}

export interface SyncStoppedPartwayOutcome {
  kind: "stoppedPartway";
  // At least 1: a send held at its first flush is stoppedBeforeSending, since nothing in it landed.
  landedFlushes: number;
  // The dry run's flush count, from its held file; absent when the send had none to read.
  flushCount?: number;
  heldBy: SyncHeldBy;
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
  ({ kind: "send" } & SyncSendLine) | SyncCreatedColumn;

export interface SyncCreatedColumn {
  kind: "createdColumn";
  table: string;
  header: string;
  column: string;
  columnId: string;
  columnType: string;
  position: string;
}

export interface SyncOutsideReference {
  cell: string;
  formula: string;
}

export type SyncCheck =
  | {
      kind: "outsideReferences";
      subject: SyncSubject;
      effect: SyncOutsideReferenceEffect;
      references: SyncOutsideReference[];
    }
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

export type SyncRefusal =
  // An entry on a Table without API access that this run doesn't grant it.
  | { kind: "noApiAccess"; table: string; path: string }
  // Any other entry on a Table this run grants API access.
  | { kind: "besideGrant"; table: string; path: string }
  // A `tables` or `tabs` entry naming one the previous ID-name map doesn't hold.
  | { kind: "noSuchTable"; entry: "tables" | "tabs"; name: string }
  // `table` names the tab for a tab's trait; `path` is the trait's, with its column in front for a column's.
  | { kind: "notSendableYet"; table: string; path: string }
  | { kind: "noSuchRule"; table: string; path: string; handles: string[] }
  // A brace naming a Table or column the previous ID-name map lacks; `columns` is braceTable's, undefined when the map holds no such Table.
  | {
      kind: "unknownBrace";
      table: string;
      path: string;
      brace: string;
      braceTable: string;
      columns: string[] | undefined;
    }
  // A `${…}` in a formula that reads as no brace, such as `@` before a Table key.
  | { kind: "malformedBrace"; table: string; path: string; brace: string }
  | {
      kind: "touchesDeletion";
      deleted: SyncSubject;
      deletedAt: string;
      touchedBy: string;
    };

export interface ForcedTrait {
  table: string;
  column: string;
  trait: string;
}

export type SyncStop = SyncConflict | SyncHeadRowsStop;

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

export interface SyncHeadRowsStop {
  kind: "headRows";
  table: string;
  tab: string;
  headerRow: number;
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
  // A Table the previous configs hold that a person deleted from the sheet.
  | { kind: "tableGone"; table: string }
  | { kind: "text"; text: string };

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

export interface SyncMovedName {
  table: string;
  column: string;
  wasId: string;
  nowId: string;
}

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
