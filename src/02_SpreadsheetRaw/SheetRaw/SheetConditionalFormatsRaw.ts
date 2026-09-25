import {
  type ConditionalFormatDeclaration,
  type ConditionalFormatRule,
  conditionalFormatRulesEqual,
  rangeEqual,
} from "../../00_Source/RawSource/ConditionalFormat";
import type { GridRangeProps } from "../../00_Source/RawSource/RawSource";
import { SheetCommonRaw } from "../ClassBases/SheetCommonRaw";
import { SheetRaw } from "../SheetRaw";
import { SpreadsheetRaw } from "../SpreadsheetRaw";

export class SheetConditionalFormatsRaw extends SheetCommonRaw {
  get ss(): SpreadsheetRaw {
    return new SpreadsheetRaw(this.spreadsheetRawProps);
  }
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
  addConditionalFormatRule(declaration: ConditionalFormatDeclaration): void {
    this.addConditionalFormatRuleAt(this.sheet.dataGridRange, declaration);
  }
  removeConditionalFormatRules(): void {
    this.removeConditionalFormatRulesAt(this.sheet.dataGridRange);
  }
  addConditionalFormatRuleAt(
    range: GridRangeProps,
    declaration: ConditionalFormatDeclaration,
  ): void {
    this.activeTable.assertRowIndexesNotStale();
    this.assertConditionalFormatIndexesNotStale();
    const rule: Extract<ConditionalFormatRule, { kind: "boolean" }> = {
      kind: "boolean",
      ranges: [range],
      condition: declaration.condition,
      format: declaration.format,
    };
    if (
      this._pendingConditionalFormatRules().some((pending) =>
        conditionalFormatRulesEqual(pending, rule),
      )
    ) {
      return;
    }
    this.updateRequests.addConditionalFormat.push({
      kind: "addConditionalFormatRule",
      index: 0,
      rule,
    });
  }
  private _pendingConditionalFormatRules(): ConditionalFormatRule[] {
    const fetched = this.sheetState.working.conditionalFormats.rules;
    const rules = fetched === undefined ? [] : [...fetched];
    const deletes = [...this.updateRequests.deleteConditionalFormat]
      .filter((operation) => operation.sheetId === this.sheetGid)
      .sort((left, right) => right.index - left.index);
    deletes.forEach((operation) => {
      rules.splice(operation.index, 1);
    });
    this.updateRequests.addConditionalFormat.forEach((operation) => {
      if (operation.rule.ranges[0]?.sheetId !== this.sheetGid) return;
      rules.splice(operation.index, 0, operation.rule);
    });
    return rules;
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
  private _removeRulesWhere(
    matches: (rule: ConditionalFormatRule) => boolean,
  ): void {
    this.activeTable.assertRowIndexesNotStale();
    this.assertConditionalFormatIndexesNotStale();
    this.conditionalFormatRules().forEach((existing, index) => {
      if (!matches(existing)) return;
      this.updateRequests.deleteConditionalFormat.push({
        kind: "deleteConditionalFormatRule",
        sheetId: this.sheetGid,
        index,
      });
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
