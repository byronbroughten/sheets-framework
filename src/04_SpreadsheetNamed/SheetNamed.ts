import type { ConditionalFormatRule } from "../00_Source/RawSource/ConditionalFormat";
import type {
  EditProtection,
  WholeSheetEditLockDeclaration,
  WholeSheetEditWarningDeclaration,
} from "../00_Source/RawSource/EditProtection";
import { SpreadsheetSchema } from "../01_SpreadsheetSchema/SpreadsheetSchema";
import type { TableName } from "../01_SpreadsheetSchema/tableConfigsTypes";
import { SheetRaw } from "../02_SpreadsheetRaw/SheetRaw";
import {
  SpreadsheetBaseNamed,
  type SpreadsheetNamedProps,
} from "./ClassBases/SpreadsheetBaseNamed";

export interface SheetNamedProps extends SpreadsheetNamedProps {
  sheetGid: number;
}

export class SheetNamed extends SpreadsheetBaseNamed {
  readonly sheetGid: number;
  constructor({ sheetGid, ...props }: SheetNamedProps) {
    super(props);
    this.sheetGid = sheetGid;
  }
  get schema(): SpreadsheetSchema {
    return new SpreadsheetSchema();
  }
  get raw(): SheetRaw {
    return new SheetRaw({
      ...this.spreadsheetRawProps,
      sheetGid: this.sheetGid,
    });
  }
  get title(): string {
    return this.raw.title;
  }
  updateTitle(title: string): this {
    this.raw.updateTitle(title);
    return this;
  }
  get tableNames(): TableName[] {
    return this.schema
      .tablesOnGid(this.sheetGid)
      .map((table) => table.tableName);
  }
  prepFetchConditionalFormatRules(): this {
    this.raw.gatherFetchConditionalFormatRules();
    return this;
  }
  conditionalFormatRules(): ConditionalFormatRule[] {
    return this.raw.conditionalFormatRules();
  }
  removeConditionalFormatRule(rule: ConditionalFormatRule): this {
    this.raw.removeConditionalFormatRule(rule);
    return this;
  }
  prepFetchEditProtections(): this {
    this.raw.gatherFetchEditProtections();
    return this;
  }
  editProtections(): EditProtection[] {
    return this.raw.editProtections();
  }
  addEditWarningWholeSheet(
    declaration: WholeSheetEditWarningDeclaration = {},
  ): this {
    this.raw.addEditWarningWholeSheet(declaration);
    return this;
  }
  addEditLockWholeSheet(declaration: WholeSheetEditLockDeclaration = {}): this {
    this.raw.addEditLockWholeSheet(declaration);
    return this;
  }
  removeEditProtection(protection: EditProtection): this {
    this.raw.removeEditProtection(protection);
    return this;
  }
  removeEditProtectionByDescription(description: string): this {
    this.raw.removeEditProtectionByDescription(description);
    return this;
  }
  removeEditProtectionById(protectionId: number): this {
    this.raw.removeEditProtectionById(protectionId);
    return this;
  }
}
