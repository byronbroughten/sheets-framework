import type { ColumnTypeBaseline } from "../../columnBaseline";
import type { ConfigSetBase } from "../configRegister";
import type { ColumnConfigStored, TableConfigStored } from "../makeConfigs";
import {
  type Brace,
  braceMessage,
  formulaPatterns,
  type FormulaText,
  noColumnMessage,
  replaceBraces,
} from "./braces";

export type BraceConfigs = Pick<
  ConfigSetBase,
  "tableConfigs" | "columnConfigs"
>;

// Only the names a compile reads, so a Table this run creates, which has no config yet, can be compiled against.
export interface BraceNames {
  tableConfigs: Record<string, Pick<TableConfigStored, "tableName">>;
  columnConfigs: Record<
    string,
    Record<string, Pick<ColumnConfigStored, "header">>
  >;
}

interface BraceFormulaProps extends FormulaText {
  configs: BraceConfigs;
}
interface BraceCompileProps extends FormulaText {
  configs: BraceNames;
}

interface LiveTable {
  tableKey: string;
  columnNamesByHeader: Map<string, string>;
}

export interface ConfiguredColumn {
  tableKey: string;
  columnName: string;
}

interface CompilableColumn {
  formula?: string;
  columnType?: ColumnTypeBaseline;
}

interface ColumnCompileProps<CL extends CompilableColumn> {
  configs: BraceNames;
  tableKey: string;
  column: CL;
}

interface ColumnTypeCompileProps {
  configs: BraceNames;
  tableKey: string;
  columnType: ColumnTypeBaseline;
}

export interface ColumnTypeProps {
  configs: BraceConfigs;
  tableKey: string;
  columnType: ColumnTypeBaseline | undefined;
}

interface BraceTextProps {
  ownTableKey: string;
  column: ConfiguredColumn;
  isThisRow: boolean;
}

export const cellFormula = {
  compileForCell({ configs, tableKey, formula }: BraceCompileProps): string {
    return replaceBraces({ tableKey, formula }, (brace) => {
      const reference = liveReferenceOf(configs, brace, formula);
      if (!brace.isThisRow) return reference;
      return `SINGLE(${reference})`;
    });
  },
  // A reference to no configured column stays as Sheets wrote it; a brace stays too, so text already in brace form comes back unchanged.
  decompile({ configs, tableKey, formula }: BraceFormulaProps): string {
    const liveTables = liveTablesByName(configs);
    const { stringLiteral, anyBrace, liveReference } = formulaPatterns;
    const pattern = new RegExp(
      `(${stringLiteral}|${anyBrace})|SINGLE\\(\\s*${liveReference}\\s*\\)|${liveReference}`,
      "g",
    );
    return formula.replace(
      pattern,
      (
        match: string,
        literal: string | undefined,
        singleTable: string | undefined,
        singleHeader: string | undefined,
        table: string,
        header: string,
      ) => {
        if (literal !== undefined) return literal;
        if (singleTable !== undefined && singleHeader !== undefined) {
          const column = configuredColumn(
            liveTables,
            singleTable,
            singleHeader,
          );
          if (column === undefined) return match;
          return braceText({ ownTableKey: tableKey, column, isThisRow: true });
        }
        const column = configuredColumn(liveTables, table, header);
        if (column === undefined) return match;
        return braceText({ ownTableKey: tableKey, column, isThisRow: false });
      },
    );
  },
  // Only a rule value starting with `=` is a formula; a list item such as `Size[cm]` stays text.
  decompileColumnType({
    configs,
    tableKey,
    columnType,
  }: ColumnTypeProps): ColumnTypeBaseline | undefined {
    if (columnType === undefined) return undefined;
    return withRuleFormulas(columnType, (formula) =>
      cellFormula.decompile({ configs, tableKey, formula }),
    );
  },
  // The inverse of decompileColumnType, for the names on the sheet when the rule lands.
  compileColumnType({
    configs,
    tableKey,
    columnType,
  }: ColumnTypeCompileProps): ColumnTypeBaseline {
    return withRuleFormulas(columnType, (formula) =>
      cellFormula.compileForCell({ configs, tableKey, formula }),
    );
  },
  // A column's formula and rule together, for a Table this run creates.
  compileColumn<CL extends CompilableColumn>({
    configs,
    tableKey,
    column,
  }: ColumnCompileProps<CL>): CL {
    const { formula, columnType } = column;
    return {
      ...column,
      ...(formula === undefined
        ? {}
        : {
            formula: cellFormula.compileForCell({
              configs,
              tableKey,
              formula,
            }),
          }),
      ...(columnType === undefined
        ? {}
        : {
            columnType: cellFormula.compileColumnType({
              configs,
              tableKey,
              columnType,
            }),
          }),
    };
  },
};

// Rewrites each rule value starting with `=`, the only ones that are formulas.
function withRuleFormulas(
  columnType: ColumnTypeBaseline,
  rewrite: (formula: string) => string,
): ColumnTypeBaseline {
  const { rule } = columnType;
  if (rule === null) return columnType;
  return {
    ...columnType,
    rule: {
      ...rule,
      values: rule.values.map((value) => {
        if (!value.startsWith("=")) return value;
        return rewrite(value);
      }),
    },
  };
}

// Bare for the formula's own Table, qualified for another's.
function braceText({ ownTableKey, column, isThisRow }: BraceTextProps): string {
  const { tableKey, columnName } = column;
  const columnText = isThisRow ? `@${columnName}` : columnName;
  if (tableKey === ownTableKey) return `\${${columnText}}`;
  return `\${${tableKey}[${columnText}]}`;
}

function liveReferenceOf(
  configs: BraceNames,
  brace: Brace,
  formula: string,
): string {
  const tableConfig = configs.tableConfigs[brace.tableKey];
  if (tableConfig === undefined) {
    throw new Error(braceMessage(brace, "names no Table", formula));
  }
  const columnConfig =
    configs.columnConfigs[brace.tableKey]?.[brace.columnName];
  if (columnConfig === undefined) {
    throw new Error(noColumnMessage(brace, formula));
  }
  return `${tableConfig.tableName}[${columnConfig.header}]`;
}

function liveTablesByName(configs: BraceConfigs): Map<string, LiveTable> {
  return Object.entries(configs.tableConfigs).reduce(
    (byName, [tableKey, tableConfig]) => {
      const columnNamesByHeader = Object.entries(
        configs.columnConfigs[tableKey] ?? {},
      ).reduce((byHeader, [columnName, columnConfig]) => {
        byHeader.set(columnConfig.header, columnName);
        return byHeader;
      }, new Map<string, string>());
      byName.set(tableConfig.tableName, { tableKey, columnNamesByHeader });
      return byName;
    },
    new Map<string, LiveTable>(),
  );
}

function configuredColumn(
  liveTables: Map<string, LiveTable>,
  liveTableName: string,
  header: string,
): ConfiguredColumn | undefined {
  const liveTable = liveTables.get(liveTableName);
  const columnName = liveTable?.columnNamesByHeader.get(header);
  if (liveTable === undefined || columnName === undefined) return undefined;
  return { tableKey: liveTable.tableKey, columnName };
}
