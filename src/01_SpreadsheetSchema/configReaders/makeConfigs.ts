import { idPrefixes } from "../idPrefixes";
import type { Value, ValueName } from "./valueSchemas";

export function makeImportLine(
  configMagerName:
    "makeTableConfigs" | "makeColumnConfigs" | "makeValueConfigs",
  makeConfigsImport: string,
): string {
  return `import { ${configMagerName} } from ${JSON.stringify(makeConfigsImport)};`;
}

function makeStructuredConfig<ST, const CF extends ST>(
  _structure: ST,
  t: CF,
): CF {
  return t;
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

export type ValueConfigsBase = Record<string, readonly string[]>;
export function makeValueConfigs<const VC extends ValueConfigsBase>(
  valueConfigs: VC,
): VC {
  return makeStructuredConfig(
    {} as Record<string, readonly string[]>,
    valueConfigs,
  );
}

interface ColumnConfigLiteral {
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
