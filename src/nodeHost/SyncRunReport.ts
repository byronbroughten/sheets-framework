import type {
  SyncCheck,
  SyncConfigs,
  SyncConflict,
  SyncDryRunResult,
  SyncHeld,
  SyncHeldGroup,
  SyncLandedItem,
  SyncNotice,
  SyncPlanEffect,
  SyncPlanItem,
  SyncRefusal,
  SyncRunResult,
  SyncSendLine,
  SyncSendResult,
  SyncStop,
  SyncStoppedPartwayOutcome,
  SyncSubject,
} from "./syncRunResult";

const layout = {
  rowLabelWidth: 30,
  wayOutWidth: 6,
  cellWidth: 16,
} as const;

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
    return [
      `sync · ${result.target} · ${result.mode}`,
      resultLine(result),
      ...section(checkTitle, result.checks.flatMap(checkLines)),
      ...section("Refused", refusalBlocks(result.refusals)),
      ...section("Stops", stopBlocks(result.stops)),
      ...sendSections(result),
      ...heldSection(result),
      ...section(planTitle(result), planLines(result.plan)),
      ...section(
        "Notices: nothing to do unless your task needs it",
        result.notices.map(noticeLine),
      ),
      ...section("Configs", configsLines(result)),
      ...section("Next", nextLines(result)),
    ];
  }
}

const checkTitle = "Check before sending: harm the gate doesn't hold";

function section(title: string, lines: string[]): string[] {
  if (lines.length === 0) return [];
  return ["", `== ${title} ==`, ...lines];
}

function indent(depth: number, lines: string[]): string[] {
  return lines.map((line) => `${" ".repeat(depth)}${line}`);
}

function counted(count: number, one: string, many: string): string {
  return `${count.toLocaleString("en-US")} ${count === 1 ? one : many}`;
}

function subjectLabel(subject: SyncSubject): string {
  if ("tab" in subject) {
    const gid = subject.gid === undefined ? "" : ` (gid ${subject.gid})`;
    return `tab ${subject.tab}${gid}`;
  }
  const { table, column, rewrittenFrom, path } = subject;
  const columnLabel =
    column !== undefined && rewrittenFrom !== undefined
      ? `${column} (you wrote ${rewrittenFrom.wrote}; ${rewrittenFrom.columnId})`
      : column;
  return [table, columnLabel, path]
    .filter((part) => part !== undefined)
    .join(" › ");
}

function heldCount(held: SyncHeld | undefined): number {
  return (held?.groups ?? []).reduce(
    (count, group) => count + (group.cells ?? 1),
    0,
  );
}

function resultLine(result: SyncRunResult): string {
  if (result.outcome.kind === "stopped") return stoppedLine(result);
  if (result.outcome.kind === "refused") return refusedLine(result);
  if (result.mode === "dry run") return readyLine(result);
  const { outcome } = result;
  if (outcome.kind === "sent") {
    return `SENT: ${outcome.summary}${checkClause(result)}.`;
  }
  return stoppedPartwayLine(outcome);
}

function checkClause({ checks }: SyncRunResult): string {
  if (checks.length === 0) return "";
  return `; ${counted(checks.length, "thing", "things")} to check (see Check before sending)`;
}

function stoppedLine(result: SyncRunResult): string {
  const { stops } = result;
  const conflicts = stops.filter(({ kind }) => kind === "conflict").length;
  const headRows = stops.length - conflicts;
  const why = [
    ...(conflicts > 0 ? [counted(conflicts, "conflict", "conflicts")] : []),
    ...(headRows > 0
      ? [`${counted(headRows, "Table", "Tables")} can't take head rows`]
      : []),
  ].join(" and ");
  const until = stops.length === 1 ? "it is resolved" : "they are resolved";
  return `STOPPED: ${why}${checkClause(result)}. Nothing was sent, and nothing will be until ${until}.`;
}

function refusedLine({ refusals }: SyncRunResult): string {
  const entries = counted(refusals.length, "entry", "entries");
  return `REFUSED BEFORE READING: ${entries} can't be planned. Nothing was read or sent; alterations.ts is unchanged.`;
}

function readyLine(result: SyncDryRunResult): string {
  const dryRun = "This was a dry run: nothing was sent.";
  if (result.held === undefined) {
    return `READY: nothing held${checkClause(result)}. ${dryRun}`;
  }
  const held = heldCount(result.held).toLocaleString("en-US");
  return `READY, ${held} HELD: --send sends only with --confirm (see Held)${checkClause(result)}. ${dryRun}`;
}

function stoppedPartwayLine({
  landedFlushes,
  flushCount,
  heldBy,
}: SyncStoppedPartwayOutcome): string {
  const refused = landedFlushes + 1;
  const landed = flushSpan(landedFlushes, flushCount);
  return `STOPPED PARTWAY: ${landed}; flush ${refused} was refused, because ${heldByReason[heldBy]}. Nothing after it was sent.`;
}

function flushSpan(landedFlushes: number, flushCount: number): string {
  if (landedFlushes === 0) return `no flush of ${flushCount} landed`;
  if (landedFlushes === 1) return `flush 1 of ${flushCount} landed`;
  return `flushes 1 to ${landedFlushes} of ${flushCount} landed`;
}

const heldByReason: Record<SyncStoppedPartwayOutcome["heldBy"], string> = {
  uncovered: "it holds items and the send had no --confirm",
  mismatched:
    "the sheet or the plan changed after the dry run, so --confirm doesn't cover its held items",
};

function checkLines(check: SyncCheck): string[] {
  if (check.kind === "outsideReferences") {
    const { subject, effect, references } = check;
    const count = references.length;
    const reads = count === 1 ? "reads" : "read";
    const where = effect === "blank" ? "outside this Table" : "on another tab";
    const after =
      effect === "blank"
        ? `${count === 1 ? "it reads" : "they read"} blank cells, with no error`
        : `${count === 1 ? "it shows" : "they show"} #REF!`;
    return [
      `${subjectLabel(subject)}: ${counted(count, "formula", "formulas")} ${where} ${reads} it. After the deletion ${after}:`,
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
  return refusals.flatMap((refusal, index) => [
    `REFUSED ${index + 1} · ${refusal.table} › ${refusal.path}`,
    ...indent(2, refusalText(refusal)),
  ]);
}

function refusalText(refusal: SyncRefusal): string[] {
  const { table } = refusal;
  if (refusal.kind === "noApiAccess") {
    return [
      `${table} has no API access yet, so sync has nothing to compare its columns against.`,
      `Granting access takes its own run: keep letApiAccess: true on ${table}, take this entry out, and run sync.`,
      "Other Tables' entries may share that run.",
      `Put this entry back in a later run, once ${table}'s columns are in idNameMap.ts.`,
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
  const { subject, label, cause, yours, agreed, live, claimedBy } = conflict;
  const yoursLabel =
    claimedBy === undefined ? "alterations.ts (yours)" : `claim (${claimedBy})`;
  const rows = [
    [yoursLabel, yours],
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

// Force's flag is never printed: an agent looks it up only after the developer's yes.
function waysOut({ yours, live, claimedBy }: SyncConflict): [string, string][] {
  if (claimedBy !== undefined) {
    return [
      [
        "adopt",
        `Set the claim to the sheet's text. Free; check first that ${claimedBy} still works on it.`,
      ],
      [
        "drop",
        `Drop the claim from ${claimedBy}. Free; the sheet owns this formula from then on and nothing guards it.`,
      ],
      ["force", `Send the claim's text over the sheet's. ${forceNeeds}`],
      ["ask", "Change nothing and ask the developer."],
    ];
  }
  return [
    [
      "adopt",
      `Remove the entry, or set it to ${live}. Free, but your change to this trait is given up.`,
    ],
    ["force", `Send ${yours} over the sheet's ${live}. ${forceNeeds}`],
    ["ask", "Change nothing and ask the developer which value they want."],
  ];
}

function sendSections(result: SyncRunResult): string[] {
  if (result.mode === "dry run") return [];
  return [
    ...section("Landed: removed from alterations.ts", landedLines(result)),
    ...section(
      "Not sent: still in alterations.ts",
      result.notSent.map(sendLine),
    ),
    ...section(
      "Didn't land: still in alterations.ts",
      result.didntLand.map(sendLine),
    ),
  ];
}

function landedLines({ landed, notSent, didntLand }: SyncSendResult): string[] {
  const lines = landed.map(landedLine);
  const isAlterationsEmpty =
    lines.length > 0 && notSent.length === 0 && didntLand.length === 0;
  if (isAlterationsEmpty) return [...lines, "alterations.ts is now empty."];
  return lines;
}

function landedLine(item: SyncLandedItem): string {
  if (item.kind === "send") return sendLine(item);
  const { table, header, column, columnId, columnType, position } = item;
  return `${table}: created column "${header}" as ${column}, ${columnId}, ${columnType}, ${position}.`;
}

function sendLine({ subject, what }: SyncSendLine): string {
  return `${subjectLabel(subject)}: ${what}`;
}

function heldSection(result: SyncRunResult): string[] {
  if (result.mode === "send" || result.held === undefined) return [];
  const { held } = result;
  return section(heldTitle(held), [
    ...held.groups.flatMap(heldGroupLines),
    `Full list, every cell with its value: ${held.fullListPath}`,
    `Token ${held.token}: binds this target and exactly this list.`,
    `Rerun with --send --confirm ${held.token} to send these.`,
  ]);
}

function heldTitle(held: SyncHeld): string {
  const confirm = "sent only with --confirm";
  const cellCounts = held.groups.flatMap(({ cells }) =>
    cells === undefined ? [] : [cells],
  );
  const count = heldCount(held);
  if (cellCounts.length < held.groups.length) {
    return `Held: ${counted(count, "item", "items")}, ${confirm}`;
  }
  const sum =
    cellCounts.length > 1
      ? `: ${cellCounts.map((cells) => cells.toLocaleString("en-US")).join(" + ")}`
      : "";
  return `Held: ${counted(count, "cell", "cells")}${sum}, ${confirm}`;
}

function heldGroupLines({ subject, what, lines }: SyncHeldGroup): string[] {
  return [`${subjectLabel(subject)}  ${what}`, ...indent(2, lines)];
}

function planTitle(result: SyncRunResult): string {
  if (result.mode === "send") return "Plan: also sent, in order";
  if (result.outcome.kind === "stopped") {
    return "Plan: what --send would send once nothing stops it";
  }
  return "Plan: what --send sends, in order";
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
    const rules = counted(
      effect.count,
      "conditional-format rule",
      "conditional-format rules",
    );
    return `Rewrites ${rules} over ${effect.table} after the ${effect.after}, since Sheets leaves them wrong.`;
  }
  if (effect.kind === "ruleRespanned") {
    return `Re-spans 1 rule that ended at row ${effect.fromRow} to the Table's last body row, ${effect.toRow}: ${effect.handle}.`;
  }
  if (effect.kind === "rulesRemoved") {
    const rules = counted(
      effect.handles.length,
      "conditional-format rule",
      "conditional-format rules",
    );
    return `Removes ${rules} wholly over it: ${effect.handles.join(", ")}.`;
  }
  if (effect.kind === "rulePartlyOver") {
    return `Leaves 1 rule partly over it, on ${effect.tab} ${effect.range} (${effect.formula}); Sheets keeps it over the blank cells.`;
  }
  if (effect.kind === "outsideReferences") {
    const formulas = counted(effect.count, "formula", "formulas");
    const after = effect.effect === "blank" ? "read blank cells" : "show #REF!";
    const reads = effect.count === 1 ? "reads" : "read";
    return `${formulas} elsewhere ${reads} it and will ${after} after the deletion (see Check before sending).`;
  }
  return `Overwrites ${counted(effect.cells, "cell", "cells")} of entered data (held, above); the other ${effect.blank.toLocaleString("en-US")} are blank.`;
}

function noticeLine(notice: SyncNotice): string {
  if (notice.kind === "text") return notice.text;
  const { table } = notice;
  return `${table}: ID is the first column. Sheets' Table name chip covers the cell above it, so ${table}'s Fill Row IDs checkbox can't be clicked. Moving ID out of first place fixes it.`;
}

// A refusal reads nothing, so there is nothing to refresh the configs from.
function configsLines(result: SyncRunResult): string[] {
  if (result.outcome.kind === "refused") return [];
  const { configs, generatedDir } = result;
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

function nextLines(result: SyncRunResult): string[] {
  const { outcome, command } = result;
  if (outcome.kind === "refused") {
    return [`Fix each refused entry, then rerun ${command}.`];
  }
  if (outcome.kind === "stopped") return [resolveLine(result.stops, command)];
  if (result.mode === "dry run") return [readyNextLine(result)];
  if (outcome.kind === "stoppedPartway") {
    return [`Rerun the dry run: ${command}`];
  }
  return [
    `sync wrote ${result.generatedDir} and ${result.alterationsPath}. Run npm run check, then commit them with your change.`,
  ];
}

function resolveLine(stops: SyncStop[], command: string): string {
  const [stop] = stops;
  if (stops.length !== 1 || stop === undefined) {
    return `Resolve each stop above, then rerun ${command}.`;
  }
  const name = stop.kind === "conflict" ? "conflict" : "stop";
  return `Resolve ${name} 1, then rerun ${command}.`;
}

function readyNextLine(result: SyncDryRunResult): string {
  const { held, checks, command } = result;
  if (held !== undefined) {
    const items = heldCount(held) === 1 ? "item" : "items";
    return `--send refuses this plan without --confirm ${held.token}, because of the held ${items}.`;
  }
  if (checks.length > 0) {
    return `Nothing held, but read Check before sending first: ${command} -- --send sends this plan.`;
  }
  return `Nothing held: ${command} -- --send sends this plan.`;
}
