import { sendableTraits } from "../01_SpreadsheetSchema/configReaders/SyncEntryRefuser";
import {
  fillRowIdsName,
  type SyncCheck,
  type SyncConfigs,
  type SyncConflict,
  type SyncDryRunOutcome,
  type SyncHeld,
  type SyncHeldBy,
  type SyncHeldGroup,
  type SyncHeldStoppedOutcome,
  type SyncLandedItem,
  type SyncNotice,
  type SyncOutsideReferenceEffect,
  type SyncPlanEffect,
  type SyncPlanItem,
  type SyncRefusal,
  type SyncRunResult,
  type SyncSendLine,
  type SyncSendOutcome,
  type SyncSentOutcome,
  type SyncStop,
  type SyncSubject,
} from "../01_SpreadsheetSchema/configReaders/syncRunResult";

const layout = {
  rowLabelWidth: 30,
  wayOutWidth: 6,
  cellWidth: 16,
} as const;

const sectionNames = {
  check: "Check before sending",
  held: "Held",
} as const;

const planTitles = {
  stopped: "Plan: what --send would send once nothing stops it",
  dryRun: "Plan: what --send sends, in order",
  send: "Plan: also sent, in order",
} as const;

// The lines that follow from how the run ended, so one dispatch over the outcome supplies them all.
interface OutcomeText {
  resultLine: string;
  planTitle: string;
  next: string;
}

export interface SyncRunReportProps {
  result: SyncRunResult;
}

// Plain text for an agent, in the shape sheets-framework#216 tested: one result line, then sections in a fixed order.
export class SyncRunReport {
  private result: SyncRunResult;
  constructor({ result }: SyncRunReportProps) {
    this.result = result;
  }
  static init(result: SyncRunResult): SyncRunReport {
    return new SyncRunReport({ result });
  }
  get text(): string {
    return this.lines.join("\n");
  }
  get lines(): string[] {
    const { result } = this;
    const outcomeText = this._outcomeText;
    return [
      `sync · ${result.target} · ${result.mode}`,
      outcomeText.resultLine,
      ...section(
        `${sectionNames.check}: harm the gate doesn't hold`,
        result.checks.flatMap(checkLines),
      ),
      ...section("Refused", refusalBlocks(result.refusals)),
      ...section("Stops", stopBlocks(result.stops)),
      ...this._sendSections,
      ...this._heldSection,
      ...section(outcomeText.planTitle, planLines(result.plan)),
      ...section(
        "Notices: nothing to do unless your task needs it",
        result.notices.map(noticeLine),
      ),
      ...section("Configs", this._configsLines),
      ...section("Next", [outcomeText.next]),
    ];
  }
  private get _outcomeText(): OutcomeText {
    const { result } = this;
    if (result.mode === "dry run") {
      return this._dryRunText(result.outcome, result.held);
    }
    return this._sendText(result.outcome);
  }
  private _dryRunText(
    outcome: SyncDryRunOutcome,
    held: SyncHeld | undefined,
  ): OutcomeText {
    if (outcome.kind === "stopped") return this._stoppedText;
    if (outcome.kind === "refused") return this._refusedText;
    return this._readyText(held);
  }
  private _sendText(outcome: SyncSendOutcome): OutcomeText {
    if (outcome.kind === "stopped") return this._stoppedText;
    if (outcome.kind === "refused") return this._refusedText;
    if (outcome.kind === "sent") return this._sentText(outcome);
    return this._heldStoppedText(outcome);
  }
  private get _stoppedText(): OutcomeText {
    const { stops, command } = this.result;
    const conflicts = stops.filter(({ kind }) => kind === "conflict").length;
    const headRowStops = stops.length - conflicts;
    const why = [
      ...(conflicts > 0 ? [counted(conflicts, "conflict", "conflicts")] : []),
      ...(headRowStops > 0
        ? [`${counted(headRowStops, "Table", "Tables")} can't take head rows`]
        : []),
    ].join(" and ");
    const until = stops.length === 1 ? "it is resolved" : "they are resolved";
    return {
      resultLine: `STOPPED: ${why}${this._checkClause}. Nothing was sent, and nothing will be until ${until}.`,
      planTitle: planTitles.stopped,
      next: resolveLine(stops, command),
    };
  }
  private get _refusedText(): OutcomeText {
    const { refusals, command } = this.result;
    const entries = counted(refusals.length, "entry", "entries");
    return {
      resultLine: `REFUSED BEFORE READING: ${entries} can't be planned. Nothing was read or sent; alterations.ts is unchanged.`,
      planTitle: planTitles.dryRun,
      next: `Fix each refused entry, then rerun ${command}.`,
    };
  }
  private _readyText(held: SyncHeld | undefined): OutcomeText {
    const { checks, command } = this.result;
    const dryRun = "This was a dry run: nothing was sent.";
    const planTitle = planTitles.dryRun;
    if (held !== undefined) {
      const count = heldCount(held);
      const items = count === 1 ? "item" : "items";
      return {
        resultLine: `READY, ${formatCount(count)} HELD: --send sends only with --confirm (see ${sectionNames.held})${this._checkClause}. ${dryRun}`,
        planTitle,
        next: `--send refuses this plan without --confirm ${held.token}, because of the held ${items}.`,
      };
    }
    const resultLine = `READY: nothing held${this._checkClause}. ${dryRun}`;
    const send = `${command} -- --send sends this plan.`;
    if (checks.length > 0) {
      return {
        resultLine,
        planTitle,
        next: `Nothing held, but read ${sectionNames.check} first: ${send}`,
      };
    }
    return { resultLine, planTitle, next: `Nothing held: ${send}` };
  }
  private _sentText({ summary }: SyncSentOutcome): OutcomeText {
    const { generatedDir, alterationsPath } = this.result;
    return {
      resultLine: `SENT: ${summary}${this._checkClause}.`,
      planTitle: planTitles.send,
      next: `sync wrote ${generatedDir} and ${alterationsPath}. Run npm run check, then commit them with your change.`,
    };
  }
  private _heldStoppedText(outcome: SyncHeldStoppedOutcome): OutcomeText {
    return {
      resultLine: heldRefusalLine(outcome),
      planTitle: planTitles.send,
      next: rerunDryRunLine(this.result.command),
    };
  }
  private get _checkClause(): string {
    const { checks } = this.result;
    if (checks.length === 0) return "";
    return `; ${counted(checks.length, "thing", "things")} to check (see ${sectionNames.check})`;
  }
  private get _sendSections(): string[] {
    const { result } = this;
    if (result.mode === "dry run") return [];
    const { landed, notSent, didntLand } = result;
    const landedLines = landed.map(landedLine);
    const isAlterationsEmpty =
      landedLines.length > 0 && notSent.length === 0 && didntLand.length === 0;
    if (isAlterationsEmpty) landedLines.push("alterations.ts is now empty.");
    return [
      ...section("Landed: removed from alterations.ts", landedLines),
      ...section("Not sent: still in alterations.ts", notSent.map(sendLine)),
      ...section(
        "Didn't land: still in alterations.ts",
        didntLand.map(sendLine),
      ),
    ];
  }
  private get _heldSection(): string[] {
    const { result } = this;
    if (result.mode === "send" || result.held === undefined) return [];
    return heldSection(result.held);
  }
  // A refusal reads nothing, so there is nothing to refresh the configs from.
  private get _configsLines(): string[] {
    const { outcome, configs, generatedDir } = this.result;
    if (outcome.kind === "refused") return [];
    const changes = [
      ...heldBackLines(configs),
      ...renamedLines(configs),
      ...namesMovedLines(configs, generatedDir),
      ...rewriteLines(configs),
      ...typedTraitLines(configs),
      ...configs.notes,
    ];
    if (changes.length === 0) {
      return ["Refreshed from the live sheet. No changes."];
    }
    return ["Refreshed from the live sheet.", ...changes];
  }
}

// A dry run's held items, their full list's path and the token covering them; a chore's dry run prints it too.
export function heldSection(held: SyncHeld): string[] {
  return section(heldTitle(held), [
    ...held.groups.flatMap(heldGroupLines),
    `Full list, every cell with its value: ${held.fullListPath}`,
    `Token ${held.token}: binds this target and exactly this list.`,
    `Rerun with --send --confirm ${held.token} to send these.`,
  ]);
}

// The result line of a send the gate refused, sync's or a chore's.
export function heldRefusalLine(outcome: SyncHeldStoppedOutcome): string {
  const reason = heldByReason[outcome.heldBy];
  if (outcome.kind === "stoppedBeforeSending") {
    return `STOPPED: the send was refused before any flush landed, because ${reason}. Nothing was sent.`;
  }
  const { landedFlushes, flushCount } = outcome;
  const landed = flushSpan(landedFlushes, flushCount);
  return `STOPPED PARTWAY: ${landed}; flush ${landedFlushes + 1} was refused, because ${reason}. Nothing after it was sent.`;
}

export function rerunDryRunLine(command: string): string {
  return `Rerun the dry run: ${command}`;
}

function section(title: string, lines: string[]): string[] {
  if (lines.length === 0) return [];
  return ["", `== ${title} ==`, ...lines];
}

function indent(depth: number, lines: string[]): string[] {
  return lines.map((line) => `${" ".repeat(depth)}${line}`);
}

function formatCount(count: number): string {
  return count.toLocaleString("en-US");
}

function counted(count: number, one: string, many: string): string {
  return `${formatCount(count)} ${count === 1 ? one : many}`;
}

function ruleCount(count: number): string {
  return counted(count, "conditional-format rule", "conditional-format rules");
}

function readsVerb(count: number): string {
  return count === 1 ? "reads" : "read";
}

const outsideReferenceWords: Record<
  SyncOutsideReferenceEffect,
  { where: string; afterOne: string; afterMany: string; inPlan: string }
> = {
  blank: {
    where: "outside this Table",
    afterOne: "it reads blank cells, with no error",
    afterMany: "they read blank cells, with no error",
    inPlan: "read blank cells",
  },
  ref: {
    where: "on another tab",
    afterOne: "it shows #REF!",
    afterMany: "they show #REF!",
    inPlan: "show #REF!",
  },
};

function subjectLabel(subject: SyncSubject): string {
  if ("tab" in subject) {
    const gid = subject.gid === undefined ? "" : ` (gid ${subject.gid})`;
    const trait = subject.path === undefined ? "" : ` › ${subject.path}`;
    return `tab ${subject.tab}${gid}${trait}`;
  }
  const { table, column, rewrittenFrom, path } = subject;
  let columnLabel = column;
  if (column !== undefined && rewrittenFrom !== undefined) {
    const { wrote, columnId } = rewrittenFrom;
    columnLabel = `${column} (you wrote ${wrote}; ${columnId})`;
  }
  return [table, columnLabel, path]
    .filter((part) => part !== undefined)
    .join(" › ");
}

function heldCount(held: SyncHeld): number {
  return held.groups.reduce((count, group) => count + (group.cells ?? 1), 0);
}

function flushSpan(
  landedFlushes: number,
  flushCount: number | undefined,
): string {
  const ofCount = flushCount === undefined ? "" : ` of ${flushCount}`;
  if (landedFlushes === 1) return `flush 1${ofCount} landed`;
  return `flushes 1 to ${landedFlushes}${ofCount} landed`;
}

const heldByReason: Record<SyncHeldBy, string> = {
  uncovered: "it holds items and the send had no --confirm",
  mismatched:
    "the sheet or the plan changed after the dry run, so --confirm doesn't cover its held items",
};

function resolveLine(stops: SyncStop[], command: string): string {
  const [stop] = stops;
  if (stops.length !== 1 || stop === undefined) {
    return `Resolve each stop above, then rerun ${command}.`;
  }
  const name = stop.kind === "conflict" ? "conflict" : "stop";
  return `Resolve ${name} 1, then rerun ${command}.`;
}

function checkLines(check: SyncCheck): string[] {
  if (check.kind === "outsideReferences") {
    const { subject, effect, references } = check;
    const count = references.length;
    const words = outsideReferenceWords[effect];
    const after = count === 1 ? words.afterOne : words.afterMany;
    return [
      `${subjectLabel(subject)}: ${counted(count, "formula", "formulas")} ${words.where} ${readsVerb(count)} it. After the deletion ${after}:`,
      ...indent(
        2,
        references.map(
          ({ cell, formula }) => `${cell.padEnd(layout.cellWidth)} ${formula}`,
        ),
      ),
    ];
  }
  if (check.kind === "followedSwap") {
    return [
      `${subjectLabel(check.subject)}: sync rewrote this entry across a swap, so it still names the column it named before, now headed "${check.header}". Check that column is the one your task means.`,
    ];
  }
  if (check.kind === "handDrawnRule") {
    const { table, tab, range, formula, handle, cause } = check;
    return [
      `${table}: 1 hand-drawn rule ${cause}: on tab ${tab}, ${range}, ${formula}. Its handle: ${handle}. Fix it by hand in Sheets, or remove it with a conditionalFormats.remove entry naming that handle.`,
    ];
  }
  return [check.text];
}

function refusalBlocks(refusals: SyncRefusal[]): string[] {
  return refusals.flatMap((refusal, index) => {
    const { heading, lines } = refusalWords(refusal);
    return [`REFUSED ${index + 1} · ${heading}`, ...indent(2, lines)];
  });
}

interface RefusalWords {
  heading: string;
  lines: string[];
}

function refusalWords(refusal: SyncRefusal): RefusalWords {
  if (refusal.kind === "touchesDeletion") {
    const { deleted, deletedAt, touchedBy } = refusal;
    return {
      heading: subjectLabel(deleted),
      lines: [
        `${deletedAt} deletes ${subjectLabel(deleted)}, and ${touchedBy} touches it in the same run.`,
        "Take one of the two out, then rerun.",
      ],
    };
  }
  if (refusal.kind === "noSuchTable") {
    const { entry, name } = refusal;
    const noun = entry === "tables" ? "Table" : "tab";
    return {
      heading: `${entry} › ${name}`,
      lines: [
        `no such ${noun}: adopt by removing the entry.`,
        `idNameMap.ts holds no ${noun} named ${name}, so it is gone from the sheet or went by another name at the last sync.`,
      ],
    };
  }
  return {
    heading: `${refusal.table} › ${refusal.path}`,
    lines: tableRefusalText(refusal),
  };
}

function tableRefusalText(
  refusal: Exclude<SyncRefusal, { kind: "touchesDeletion" | "noSuchTable" }>,
): string[] {
  if (refusal.kind === "notSendableYet") {
    return [
      `sync can't send ${refusal.path} yet: for now it sends only a Table's ${[...sendableTraits.table].join(", ")}, a column's ${[...sendableTraits.column].join(", ")} and a tab's ${[...sendableTraits.tab].join(", ")}.`,
      "Take this entry out, then rerun.",
    ];
  }
  const { table } = refusal;
  if (refusal.kind === "noApiAccess") {
    return [
      `${table} has no API access, so sync has nothing to compare it against: grant \`letApiAccess\` first.`,
      ...grantRunLines(
        table,
        `take this entry out, put letApiAccess: true on ${table}`,
      ),
    ];
  }
  if (refusal.kind === "besideGrant") {
    return [
      `${table} has no API access yet, so sync has nothing to compare its columns against.`,
      ...grantRunLines(
        table,
        `keep letApiAccess: true on ${table}, take this entry out`,
      ),
    ];
  }
  if (refusal.kind === "noSuchRule") {
    return [
      `No such rule: no rule on ${table} has this handle now. A person changed the rule after the last sync`,
      "(a handle comes from its columns, formula, condition and format). Adopt by removing the entry.",
      ...ruleHandleLines(table, refusal.handles),
    ];
  }
  return [
    `${refusal.brace} names no column of ${table}. Its columns: ${refusal.columns.join(", ")}.`,
  ];
}

function grantRunLines(table: string, steps: string): string[] {
  return [
    `Granting access takes its own run: ${steps}, and run sync.`,
    "Other Tables' entries may share that run.",
    `Put this entry back in a later run, once ${table}'s columns are in idNameMap.ts.`,
  ];
}

function ruleHandleLines(table: string, handles: string[]): string[] {
  if (handles.length === 0) return [`${table} has no rules now.`];
  return [`${table}'s rules now, by handle:`, ...indent(2, handles)];
}

function stopBlocks(stops: SyncStop[]): string[] {
  return stops.flatMap((stop, index) => [
    ...(index === 0 ? [] : [""]),
    ...stopBlock(stop, index + 1),
  ]);
}

function stopBlock(stop: SyncStop, number: number): string[] {
  if (stop.kind === "conflict") return conflictBlock(stop, number);
  const { table, tab, headerRow } = stop;
  return [
    `STOP ${number} · ${table} › letApiAccess`,
    ...indent(2, [
      `${table}'s header row is on row ${headerRow} of its tab, ${tab}. A Table needs its header on row 4 to take head rows`,
      `(rows 1 to 3: column IDs, action row, heading). Move ${table} so its header sits on row 4, then rerun.`,
      "The entry stays in alterations.ts.",
    ]),
  ];
}

function conflictBlock(conflict: SyncConflict, number: number): string[] {
  const { subject, label, cause, yours, agreed, live } = conflict;
  const rows = [
    ["alterations.ts (yours)", yours],
    ["configs (last agreed)", agreed],
    ["live sheet", live],
  ].map(([name = "", value = ""]) => name.padEnd(layout.rowLabelWidth) + value);
  return [
    `CONFLICT ${number} · ${subjectLabel(subject)} · ${label}`,
    ...indent(2, [...rows, cause, "Ways out:"]),
    ...indent(
      4,
      waysOut(conflict).map(
        ([way, text]) => `${way.padEnd(layout.wayOutWidth)} ${text}`,
      ),
    ),
  ];
}

const forceNeeds =
  "Needs the developer's yes for this trait; the yes covers one run.";

function waysOut(conflict: SyncConflict): [string, string][] {
  const { yours, live } = conflict;
  const adopt: [string, string] = [
    "adopt",
    `Remove the entry, or set it to ${live}. Free, but your change to this trait is given up.`,
  ];
  const ask: [string, string] = [
    "ask",
    "Change nothing and ask the developer which value they want.",
  ];
  const forceFlag = forceFlagOf(conflict);
  if (forceFlag === undefined) return [adopt, ask];
  return [
    adopt,
    [
      "force",
      `Send ${yours} over the sheet's ${live}. ${forceNeeds} ${forceFlag}`,
    ],
    ask,
  ];
}

// --force addresses <table>/<column>/<trait>, so a Table's or tab's own trait, or a deleted column's, has no force to offer.
function forceFlagOf({
  subject,
  isColumnGone,
}: SyncConflict): string | undefined {
  if (isColumnGone === true) return undefined;
  if (!("column" in subject)) return undefined;
  const { table, column, path } = subject;
  if (column === undefined || path === undefined) return undefined;
  return `Flag: --force ${table}/${column}/${path}`;
}

function landedLine(item: SyncLandedItem): string {
  if (item.kind === "send") return sendLine(item);
  const { table, header, column, columnId, columnType, position } = item;
  return `${table}: created column "${header}" as ${column}, ${columnId}, ${columnType}, ${position}.`;
}

function sendLine({ subject, what }: SyncSendLine): string {
  return `${subjectLabel(subject)}: ${what}`;
}

function heldTitle(held: SyncHeld): string {
  const confirm = "sent only with --confirm";
  const cellCounts = held.groups.flatMap(({ cells }) =>
    cells === undefined ? [] : [cells],
  );
  const count = heldCount(held);
  if (cellCounts.length < held.groups.length) {
    return `${sectionNames.held}: ${counted(count, "item", "items")}, ${confirm}`;
  }
  let sum = "";
  if (cellCounts.length > 1) {
    sum = `: ${cellCounts.map(formatCount).join(" + ")}`;
  }
  return `${sectionNames.held}: ${counted(count, "cell", "cells")}${sum}, ${confirm}`;
}

function heldGroupLines({ subject, what, lines }: SyncHeldGroup): string[] {
  const heading =
    subject === undefined ? what : `${subjectLabel(subject)}  ${what}`;
  return [heading, ...indent(2, lines)];
}

function planLines(plan: SyncPlanItem[]): string[] {
  return plan.flatMap((item, index) => [
    `${String(index + 1).padStart(2)}. ${sendLine(item)}`,
    ...indent(4, item.effects.map(effectLine)),
  ]);
}

function effectLine(effect: SyncPlanEffect): string {
  if (effect.kind === "dropdownReset") {
    const { columns, setting } = effect;
    const dropdowns = counted(columns.length, "dropdown", "dropdowns");
    const why =
      setting === "ask"
        ? 'dropdownStyleReset: "ask", so held, above'
        : 'dropdownStyleReset: "allow"';
    return `Resets the style and colours of ${dropdowns}: ${columns.join(", ")} (${why}).`;
  }
  if (effect.kind === "addedFirst") {
    const { outranks } = effect;
    if (outranks.length === 0) {
      return "Added first; it outranks no rule on these cells.";
    }
    return `Added first; it outranks ${counted(outranks.length, "rule", "rules")} on these cells: ${outranks.join(", ")}.`;
  }
  if (effect.kind === "rulesRewritten") {
    return `Rewrites ${ruleCount(effect.count)} over ${effect.table} after the ${effect.after}, since Sheets leaves them wrong.`;
  }
  if (effect.kind === "ruleRespanned") {
    return `Re-spans 1 rule that ended at row ${effect.fromRow} to the Table's last body row, ${effect.toRow}: ${effect.handle}.`;
  }
  if (effect.kind === "rulesRemoved") {
    const { handles } = effect;
    return `Removes ${ruleCount(handles.length)} wholly over it: ${handles.join(", ")}.`;
  }
  if (effect.kind === "rulePartlyOver") {
    return `Leaves 1 rule partly over it, on ${effect.tab} ${effect.range} (${effect.formula}); Sheets keeps it over the blank cells.`;
  }
  if (effect.kind === "outsideReferences") {
    const { count } = effect;
    const formulas = counted(count, "formula", "formulas");
    const after = outsideReferenceWords[effect.effect].inPlan;
    return `${formulas} elsewhere ${readsVerb(count)} it and will ${after} after the deletion (see ${sectionNames.check}).`;
  }
  return `Overwrites ${counted(effect.cells, "cell", "cells")} of entered data (held, above); the other ${formatCount(effect.blank)} are blank.`;
}

function noticeLine(notice: SyncNotice): string {
  if (notice.kind === "text") return notice.text;
  const { table } = notice;
  return `${table}: ID is the first column. Sheets' Table name chip covers the cell above it, so ${table}'s ${fillRowIdsName} checkbox can't be clicked. Moving ID out of first place fixes it.`;
}

function heldBackLines({ heldBack }: SyncConfigs): string[] {
  if (heldBack.length === 0) return [];
  return [
    "Held back while a conflict is open, so the next run still sees it:",
    ...indent(
      2,
      heldBack.map(
        ({ subject, value }) =>
          `${subjectLabel(subject)}: the configs keep ${value} for this trait until the conflict is resolved.`,
      ),
    ),
  ];
}

function renamedLines({ renamed }: SyncConfigs): string[] {
  if (renamed.length === 0) return [];
  return [
    "Renamed by a person since the last sync; the codebase name follows the header:",
    ...indent(
      2,
      renamed.map(
        ({ table, from, to, headerFrom, headerTo }) =>
          `${table} › ${from} → ${to}    header "${headerFrom}" → "${headerTo}"`,
      ),
    ),
  ];
}

function namesMovedLines(
  { namesMoved }: SyncConfigs,
  generatedDir: string,
): string[] {
  if (namesMoved.length === 0) return [];
  return [
    `Names that now point at a different column than at the last sync. Reread ${generatedDir}idNameMap.ts before writing another entry:`,
    ...indent(
      2,
      namesMoved.map(
        ({ table, column, wasId, nowId }) =>
          `${table} › ${column}    was ${wasId}, now ${nowId}`,
      ),
    ),
  ];
}

function rewriteLines({ rewrites }: SyncConfigs): string[] {
  if (rewrites.length === 0) return [];
  return [
    "Rewritten in alterations.ts, so each entry still names the column it named before:",
    ...indent(
      2,
      rewrites.map(({ from, to, why }) => {
        const reason = why === undefined ? "" : `    (${why})`;
        return `${from} → ${to}${reason}`;
      }),
    ),
  ];
}

function typedTraitLines({ typedTraitChanges }: SyncConfigs): string[] {
  if (typedTraitChanges.length === 0) return [];
  return [
    "Typed traits a person changed since the last sync, on Tables with API access. tsc fails where code relied on the old ones:",
    ...indent(
      2,
      typedTraitChanges.map(
        ({ table, column, trait, from, to }) =>
          `${table} › ${column}    ${trait} ${from} → ${to}`,
      ),
    ),
  ];
}
