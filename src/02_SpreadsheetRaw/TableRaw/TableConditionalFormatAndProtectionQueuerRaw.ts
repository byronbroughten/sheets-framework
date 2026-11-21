import type {
  ConditionalFormatDeclaration,
  ModelableConditionalFormatRule,
} from "../../00_Source/RawSource/ConditionalFormat.ts";
import type {
  EditLockDeclaration,
  EditProtectionContent,
  EditWarningDeclaration,
} from "../../00_Source/RawSource/EditProtection.ts";
import type { GridRangeProps } from "../../00_Source/RawSource/RawSource.ts";
import { SheetIndex } from "../../00_Source/RawSource/SheetIndex.ts";
import type { TableOrigin } from "../../01_SpreadsheetSchema/TableOrigin.ts";
import { TableBaseRaw } from "../ClassBases/TableBaseRaw.ts";
import type {
  TableConditionalFormatRule,
  TableEditProtection,
  TableGridRange,
} from "../ClassTypes/StateRaw.ts";
import { TableRaw } from "../TableRaw.ts";

export class TableConditionalFormatAndProtectionQueuerRaw extends TableBaseRaw {
  get table(): TableRaw {
    return new TableRaw(this.tableRawProps);
  }
  // Checked against what the sheet holds now; converted again at gathering, after this batch's shifts.
  addConditionalFormatRuleAt(
    range: TableGridRange,
    declaration: ConditionalFormatDeclaration,
  ): this {
    const { table } = this;
    table.sheet.assertRowIndexesNotStale();
    table.sheet.assertConditionalFormatIndexesNotStale();
    const rule = this._sheetConditionalFormatRule(
      { range, ...declaration },
      this.tableOrigin(),
    );
    const queued = this._queuedOnSheet((queuer) =>
      queuer._queuedSheetConditionalFormatRules(),
    );
    if (table.sheet.hasPendingConditionalFormatRule(rule, queued)) return this;
    table.queueTableWrite({
      action: "addConditionalFormatRule",
      range,
      ...declaration,
    });
    return this;
  }
  private _queuedSheetConditionalFormatRules(): ModelableConditionalFormatRule[] {
    const origin = this.tableOrigin();
    return this.table.writes.conditionalFormatRules.map((rule) =>
      this._sheetConditionalFormatRule(rule, origin),
    );
  }
  removeConditionalFormatRulesAt(range: TableGridRange): this {
    this.table.sheet.removeConditionalFormatRulesAt(
      this._sheetGridRange(range, this.tableOrigin()),
    );
    return this;
  }
  addEditWarningAt(
    range: TableGridRange,
    declaration: EditWarningDeclaration,
  ): this {
    return this._queueEditProtection({
      kind: "warning",
      range,
      description: declaration.description ?? "",
      users: [],
      groups: [],
      unprotectedRanges: [],
    });
  }
  addEditLockAt(range: TableGridRange, declaration: EditLockDeclaration): this {
    return this._queueEditProtection({
      kind: "lock",
      range,
      description: declaration.description ?? "",
      users: declaration.users ?? [],
      groups: declaration.groups ?? [],
      unprotectedRanges: [],
    });
  }
  removeEditProtectionsAt(range: TableGridRange): this {
    this.table.sheet.removeEditProtectionsAt(
      this._sheetGridRange(range, this.tableOrigin()),
    );
    return this;
  }
  gatherConditionalFormatRuleOperation(rule: TableConditionalFormatRule): void {
    const { table } = this;
    table.writeOperations.addConditionalFormatRule.push({
      kind: "addConditionalFormatRule",
      index: 0,
      rule: this._sheetConditionalFormatRule(rule, table.originAtGathering()),
    });
  }
  gatherEditProtectionOperation(protection: TableEditProtection): void {
    const { table } = this;
    let origin = this.tableOrigin();
    if (hasTableRows(protection.range)) origin = table.originAtGathering();
    table.writeOperations.addProtectedRange.push({
      kind: "addProtectedRange",
      protection: this._sheetEditProtection(protection, origin),
    });
  }
  private _queuedOnSheet<VL>(
    queuedOf: (queuer: TableConditionalFormatAndProtectionQueuerRaw) => VL[],
  ): VL[] {
    return this.table.tablesOnSheet.flatMap((other) =>
      queuedOf(other.conditionalFormatAndProtectionQueuer),
    );
  }
  private _sheetConditionalFormatRule(
    { range, condition, format }: TableConditionalFormatRule,
    origin: TableOrigin,
  ): ModelableConditionalFormatRule {
    return {
      kind: "boolean",
      ranges: [this._sheetGridRange(range, origin)],
      condition,
      format,
    };
  }
  private _sheetEditProtection(
    { range, ...protection }: TableEditProtection,
    origin: TableOrigin,
  ): EditProtectionContent {
    return { ...protection, range: this._sheetGridRange(range, origin) };
  }
  // Absent bounds stay absent, as the matching of fetched rules and protections expects.
  private _sheetGridRange(
    { startRowIndex, endRowIndex, startColIndex, endColIndex }: TableGridRange,
    origin: TableOrigin,
  ): GridRangeProps {
    const range: GridRangeProps = {
      sheetId: this.sheetGid,
      startRowIndex: SheetIndex.row(0),
    };
    if (startRowIndex !== undefined) {
      range.startRowIndex = origin.sheetRowIndex(startRowIndex);
    }
    if (endRowIndex !== undefined) {
      range.endRowIndex = origin.sheetRowIndex(endRowIndex);
    }
    if (startColIndex !== undefined) {
      range.startColumnIndex = origin.sheetColIndex(startColIndex);
    }
    if (endColIndex !== undefined) {
      range.endColumnIndex = origin.sheetColIndex(endColIndex);
    }
    return range;
  }
  private _queueEditProtection(protection: TableEditProtection): this {
    const { table } = this;
    // A same-batch shift carries only a Table whose properties have arrived.
    table.assertTableIsKnown();
    if (hasTableRows(protection.range)) table.sheet.assertRowIndexesNotStale();
    table.sheet.assertEditProtectionsNotStale();
    const queued = this._queuedOnSheet((queuer) =>
      queuer._queuedSheetEditProtections(),
    );
    if (
      table.sheet.hasPendingEditProtection(
        this._sheetEditProtection(protection, this.tableOrigin()),
        queued,
      )
    ) {
      return this;
    }
    table.queueTableWrite({ action: "addEditProtection", ...protection });
    return this;
  }
  private _queuedSheetEditProtections(): EditProtectionContent[] {
    const origin = this.tableOrigin();
    return this.table.writes.editProtections.map((protection) =>
      this._sheetEditProtection(protection, origin),
    );
  }
}

// A whole column has none, so moved rows don't bar it.
function hasTableRows(range: TableGridRange): boolean {
  return range.startRowIndex !== undefined;
}
