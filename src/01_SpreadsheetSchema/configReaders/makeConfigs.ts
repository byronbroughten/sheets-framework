import type { TabColor } from "../../00_Source/RawSource/RawSource";
import type {
  ActionRowCellBaseline,
  ColumnTypeBaseline,
} from "../columnBaseline";
import { idPrefixes } from "../idPrefixes";
import type { RuleRecordSet } from "./RuleRecords";
import type { Value, ValueName } from "./valueSchemas";

type ConfigMakerName =
  | "makeTableConfigs"
  | "makeTabConfigs"
  | "makeConditionalFormatConfigs"
  | "makeRetiredIdPrefixes"
  | "makeColumnConfigs"
  | "makeValueConfigs";

export function makeImportLine(
  configMakerNames: readonly ConfigMakerName[],
): string {
  return `import { ${configMakerNames.join(", ")} } from "@byronbroughten/sheets-framework";`;
}

function makeStructuredConfig<ST, const CF extends ST>(
  _structure: ST,
  config: CF,
): CF {
  return config;
}

export interface TableConfigStored {
  tableId: string;
  tableName: string;
  sheetGid: number;
  idPrefix: string;
  hasIdColumn: boolean;
  hasNameColumn: boolean;
}
export type TableConfigsBase = Record<string, TableConfigStored>;
export function makeTableConfigs<TC extends TableConfigsBase>(
  tableConfigs: TC,
): TC {
  assertUniqueIdPrefixes(tableConfigs);
  return tableConfigs;
}

function assertUniqueIdPrefixes(
  configs: Record<string, { idPrefix: string }>,
): void {
  idPrefixes.assertUnique(
    Object.entries(configs).map(([label, config]) => ({
      label,
      idPrefix: config.idPrefix,
    })),
  );
}

// Order is left out: a tab's place in the tab bar is never part of the baseline.
export interface TabConfigStored {
  sheetGid: number;
  title: string;
  color: TabColor | null;
  frozenRows: number;
  frozenColumns: number;
  hidden: boolean;
}
export type TabConfigsBase = Record<string, TabConfigStored>;
// Every generated data file is a literal passed through its constructor from makeConfigs, never a bare literal, so this and makeRetiredIdPrefixes check nothing.
export function makeTabConfigs<TC extends TabConfigsBase>(tabConfigs: TC): TC {
  return tabConfigs;
}

// Records under tables by Table key and under tabs by tab name, each keyed by its rule handle.
export type ConditionalFormatConfigsBase = RuleRecordSet;
export function makeConditionalFormatConfigs<
  CF extends ConditionalFormatConfigsBase,
>(conditionalFormatConfigs: CF): CF {
  return conditionalFormatConfigs;
}

export function makeRetiredIdPrefixes(
  retiredIdPrefixes: readonly string[],
): readonly string[] {
  return retiredIdPrefixes;
}

export type FixedListMember = string | number | boolean;
export type FixedList = readonly FixedListMember[];
export interface LiveReference {
  kind: "liveReference";
}
export type ValueConfigBase = FixedList | LiveReference;
export type ValueConfigsBase = Record<string, ValueConfigBase>;
export function makeValueConfigs<const VC extends ValueConfigsBase>(
  valueConfigs: VC,
): VC {
  return makeStructuredConfig({} as ValueConfigsBase, valueConfigs);
}

// Each baseline trait is optional so an older configs file still loads; absent means unrecorded, null means recorded empty.
interface ColumnBaselineLiteral {
  formula?: string | null;
  columnType?: ColumnTypeBaseline;
  actionRowCell?: ActionRowCellBaseline | null;
  groupHeading1?: string | null;
}
interface ColumnConfigLiteral extends ColumnBaselineLiteral {
  columnId: string;
  header: string;
  isFormula: boolean;
  emptyValueAllowed: boolean;
}
export interface ColumnConfigStored<
  VN extends ValueName = ValueName,
> extends ColumnConfigLiteral {
  valueName: VN;
  customDefaultValue: Value<VN> | null;
}

type TableColumnConfigs = Record<string, ColumnConfigStored>;
export type ColumnConfigsGeneric = Record<string, TableColumnConfigs>;

interface ColumnConfigLiteralStored<
  VN extends string,
> extends ColumnConfigLiteral {
  valueName: VN;
  customDefaultValue: unknown;
}
// Names no Register-derived type, so the generated literal can fill Register without a cycle.
export type ColumnConfigsBase<VN extends string = string> = Record<
  string,
  Record<string, ColumnConfigLiteralStored<VN>>
>;

// VN keeps each valueName a literal instead of widening it to string.
export function makeColumnConfigs<
  VN extends string,
  CC extends ColumnConfigsBase<VN>,
>(columnConfigs: CC): CC {
  return makeStructuredConfig({} as ColumnConfigsBase, columnConfigs);
}
