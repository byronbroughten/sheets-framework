import {
  type ConditionalFormatRule,
  conditionalFormatRulesEqual,
  type ModelableConditionalFormatRule,
} from "../../00_Source/RawSource/ConditionalFormat";
import {
  type GridRangeProps,
  rangeEqual,
} from "../../00_Source/RawSource/RawSource";
import { SheetBaseRaw } from "../ClassBases/SheetBaseRaw";
import { SheetRaw } from "../SheetRaw";

export class SheetConditionalFormatsRaw extends SheetBaseRaw {
  get sheet(): SheetRaw {
    return new SheetRaw(this.sheetRawProps);
  }
  gatherFetchConditionalFormatRules(): void {
    this.sheetState.fetchQueue.gatherConditionalFormats = true;
  }
  conditionalFormatRules(): ConditionalFormatRule[] {
    const rules = this.sheetState.working.conditionalFormats.rules;
    if (rules === undefined) {
      throw new Error(
        `Conditional format rules have not been fetched for sheetGid ${this.sheetGid}.`,
      );
    }
    return rules;
  }
  // A Table's rules are queued on the Table, so it passes the ones it holds.
  hasPendingConditionalFormatRule(
    rule: ModelableConditionalFormatRule,
    queuedOnTable: ModelableConditionalFormatRule[],
  ): boolean {
    return [...this._pendingConditionalFormatRules(), ...queuedOnTable].some(
      (pending) => conditionalFormatRulesEqual(pending, rule),
    );
  }
  private _pendingConditionalFormatRules(): ConditionalFormatRule[] {
    const fetched = this.sheetState.working.conditionalFormats.rules;
    const rules = fetched === undefined ? [] : [...fetched];
    const deletes = [...this.writeOperations.deleteConditionalFormatRule]
      .filter((operation) => operation.sheetId === this.sheetGid)
      .sort((left, right) => right.index - left.index);
    deletes.forEach((operation) => {
      rules.splice(operation.index, 1);
    });
    return rules;
  }
  updateConditionalFormatRule(
    index: number,
    rule: ModelableConditionalFormatRule,
  ): void {
    this.assertConditionalFormatIndexesNotStale();
    this._queueRuleUpdate(index, rule);
  }
  // No stale check: the caller already took off the deletes an earlier flush sent, so the index is the one at send time.
  updateConditionalFormatRuleReindexed(
    index: number,
    rule: ModelableConditionalFormatRule,
  ): void {
    this._queueRuleUpdate(index, rule);
  }
  // Index 0 needs no stale check, since it names no fetched rule; the flusher sends adds before updates.
  addConditionalFormatRuleFirst(rule: ModelableConditionalFormatRule): void {
    this.sheet.assertRowIndexesNotStale();
    this.writeOperations.addConditionalFormatRule.push({
      kind: "addConditionalFormatRule",
      index: 0,
      rule,
    });
  }
  private _queueRuleUpdate(
    index: number,
    rule: ModelableConditionalFormatRule,
  ): void {
    this.writeOperations.updateConditionalFormatRule.push({
      kind: "updateConditionalFormatRule",
      sheetId: this.sheetGid,
      index,
      rule,
    });
  }
  removeConditionalFormatRulesAt(range: GridRangeProps): void {
    this._removeRulesWhere(
      (rule) => rule.ranges.length === 1 && rangeEqual(range, rule.ranges[0]),
    );
  }
  removeConditionalFormatRule(rule: ConditionalFormatRule): void {
    this._removeRulesWhere((existing) =>
      conditionalFormatRulesEqual(existing, rule),
    );
  }
  // The one rule at its fetched index: two equal rules are two, and a match by content would take both.
  removeConditionalFormatRuleByIndex(index: number): void {
    this.assertConditionalFormatIndexesNotStale();
    if (this.conditionalFormatRules()[index] === undefined) {
      throw new Error(
        `No conditional format rule at index ${index} for sheetGid ${this.sheetGid}.`,
      );
    }
    this._queueRuleDelete(index);
  }
  private _removeRulesWhere(
    matches: (rule: ConditionalFormatRule) => boolean,
  ): void {
    this.sheet.assertRowIndexesNotStale();
    this.assertConditionalFormatIndexesNotStale();
    this.conditionalFormatRules().forEach((existing, index) => {
      if (!matches(existing)) return;
      this._queueRuleDelete(index);
    });
  }
  // Indexed as fetched: the flusher sends a sheet's deletes highest index first, so each is still right when sent.
  private _queueRuleDelete(index: number): void {
    this.writeOperations.deleteConditionalFormatRule.push({
      kind: "deleteConditionalFormatRule",
      sheetId: this.sheetGid,
      index,
    });
  }
  markConditionalFormatIndexesStale(): void {
    this.sheetState.working.conditionalFormats.isStale = true;
  }
  assertConditionalFormatIndexesNotStale(): void {
    if (!this.sheetState.working.conditionalFormats.isStale) return;
    throw new Error(
      `Conditional format indexes are stale for sheetGid ${this.sheetGid}. Re-fetch the sheet's rules before mutating them again.`,
    );
  }
  integrateConditionalFormatRules(rules: ConditionalFormatRule[]): void {
    this.sheetState.working.conditionalFormats.rules = rules;
    this.sheetState.fetchQueue.gatherConditionalFormats = false;
    this.sheetState.working.conditionalFormats.isStale = false;
  }
}
