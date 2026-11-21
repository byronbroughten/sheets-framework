import type { DryRunHeld } from "../01_SpreadsheetSchema/configReaders/dryRunHeld";
import {
  type FlushStoppedOutcome,
  isFillFailed,
} from "../01_SpreadsheetSchema/configReaders/flushStoppedOutcome";
import type {
  SyncDryRunOutcome,
  SyncRunResult,
  SyncSendOutcome,
  SyncSentOutcome,
} from "../01_SpreadsheetSchema/configReaders/syncRunResult";
import { GateReport, heldSectionName } from "./GateReport";
import { ReportText } from "./ReportText";
import { checkLines, checkSectionName } from "./SyncRunReport/checks";
import { configsLines } from "./SyncRunReport/configs";
import { noticeLines } from "./SyncRunReport/notices";
import { planLines } from "./SyncRunReport/plan";
import { refusalLines } from "./SyncRunReport/refusals";
import { sendLines } from "./SyncRunReport/sends";
import { stopLines } from "./SyncRunReport/stops";

const { section, formatCount, counted } = ReportText;

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
        `${checkSectionName}: harm the gate doesn't hold`,
        result.checks.flatMap(checkLines.of),
      ),
      ...section("Refused", refusalLines.blocks(result.refusals)),
      ...section("Stops", stopLines.blocks(result.stops)),
      ...this._sendSections,
      ...this._heldSection,
      ...section(outcomeText.planTitle, planLines.of(result.plan)),
      ...section(
        "Notices: nothing to do unless your task needs it",
        result.notices.map(noticeLines.of),
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
    held: DryRunHeld | undefined,
  ): OutcomeText {
    if (outcome.kind === "stopped") return this._stoppedText;
    if (outcome.kind === "refused") return this._refusedText;
    return this._readyText(held);
  }
  private _sendText(outcome: SyncSendOutcome): OutcomeText {
    if (outcome.kind === "stopped") return this._stoppedText;
    if (outcome.kind === "refused") return this._refusedText;
    if (outcome.kind === "sent") return this._sentText(outcome);
    return this._flushStoppedText(outcome);
  }
  private get _stoppedText(): OutcomeText {
    const { stops, command } = this.result;
    const why = stopLines.why(stops);
    const until = stops.length === 1 ? "it is resolved" : "they are resolved";
    return {
      resultLine: `STOPPED: ${why}${this._checkClause}. Nothing was sent, and nothing will be until ${until}.`,
      planTitle: planTitles.stopped,
      next: stopLines.resolveLine(stops, command),
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
  private _readyText(held: DryRunHeld | undefined): OutcomeText {
    const { checks, command } = this.result;
    const dryRun = "This was a dry run: nothing was sent.";
    const planTitle = planTitles.dryRun;
    if (held !== undefined) {
      const count = GateReport.heldCount(held);
      const items = count === 1 ? "item" : "items";
      return {
        resultLine: `READY, ${formatCount(count)} HELD: --send sends only with --confirm (see ${heldSectionName})${this._checkClause}. ${dryRun}`,
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
        next: `Nothing held, but read ${checkSectionName} first: ${send}`,
      };
    }
    return { resultLine, planTitle, next: `Nothing held: ${send}` };
  }
  private _sentText({ summary }: SyncSentOutcome): OutcomeText {
    return {
      resultLine: `SENT: ${summary}${this._checkClause}.`,
      planTitle: planTitles.send,
      next: this._wroteLine,
    };
  }
  private _flushStoppedText(outcome: FlushStoppedOutcome): OutcomeText {
    const { command } = this.result;
    let next = GateReport.rerunDryRunLine(command);
    if (isFillFailed(outcome)) {
      next = `${this._wroteLine} Then rerun the dry run to fill the blank IDs: ${command}`;
    }
    return {
      resultLine: GateReport.flushStoppedLine(outcome),
      planTitle: planTitles.send,
      next,
    };
  }
  private get _wroteLine(): string {
    const { generatedDir, alterationsPath } = this.result;
    return `sync wrote ${generatedDir} and ${alterationsPath}. Run npm run check, then commit them with your change.`;
  }
  private get _checkClause(): string {
    const { checks } = this.result;
    if (checks.length === 0) return "";
    return `; ${counted(checks.length, "thing", "things")} to check (see ${checkSectionName})`;
  }
  private get _sendSections(): string[] {
    const { result } = this;
    const metLines = result.alreadyMet.map(sendLines.of);
    if (result.mode === "dry run") {
      return section(
        "Already met: --send removes from alterations.ts",
        metLines,
      );
    }
    const { landed, notSent, didntLand, outcome, hasEmptiedAlterations } =
      result;
    const landedLines = landed.map(sendLines.landed);
    if (hasEmptiedAlterations) {
      const lastSectionLines = metLines.length > 0 ? metLines : landedLines;
      lastSectionLines.push("alterations.ts is now empty.");
    }
    // A stopped send, even one stopped partway, leaves its met alterations pending, unless only its fill failed, after the files were built.
    const isMetRemoved = outcome.kind === "sent" || isFillFailed(outcome);
    const metTitle = isMetRemoved ? "removed from" : "still in";
    return [
      ...section("Landed: removed from alterations.ts", landedLines),
      ...section(`Already met: ${metTitle} alterations.ts`, metLines),
      ...section(
        "Not sent: still in alterations.ts",
        notSent.map(sendLines.of),
      ),
      ...section(
        "Didn't land: still in alterations.ts",
        didntLand.map(sendLines.of),
      ),
    ];
  }
  private get _heldSection(): string[] {
    const { result } = this;
    if (result.mode === "send" || result.held === undefined) return [];
    return GateReport.heldSection(result.held);
  }
  // A refusal reads nothing, so there is nothing to refresh the configs from.
  private get _configsLines(): string[] {
    const { outcome, configs, generatedDir } = this.result;
    if (outcome.kind === "refused") return [];
    const changes = configsLines.changes(configs, generatedDir);
    if (changes.length === 0) {
      return ["Refreshed from the live sheet. No changes."];
    }
    return ["Refreshed from the live sheet.", ...changes];
  }
}
