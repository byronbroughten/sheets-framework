import {
  type SheetColIndex,
  SheetIndex,
  type SheetRowIndex,
} from "../../00_Source/RawSource/SheetIndex";
import { SpreadsheetBaseSchema } from "../SpreadsheetBaseSchema";
import type { ConfigSetBase } from "./configRegister";

const schema = new SpreadsheetBaseSchema();

export type BraceConfigs = Pick<
  ConfigSetBase,
  "tableConfigs" | "columnConfigs"
>;

interface FormulaText {
  tableKey: string;
  formula: string;
}
interface BraceFormulaProps extends FormulaText {
  configs: BraceConfigs;
}
interface RuleFormulaProps extends FormulaText {
  layout: RuleLayout;
}

interface Brace {
  text: string;
  tableKey: string;
  columnName: string;
  isWholeColumn: boolean;
}

export interface RuleLayout {
  headerTopLeft: { rowIndex: SheetRowIndex; colIndex: SheetColIndex };
  columnNamesLeftToRight: readonly string[];
}

interface LiveTable {
  tableKey: string;
  columnNamesByHeader: Map<string, string>;
}

interface ConfiguredColumn {
  tableKey: string;
  columnName: string;
}

const formulaPatterns = {
  stringLiteral: /"(?:[^"]|"")*"/.source,
  brace: /\$\{([^{}[\]"]+?)(?:\[([^{}[\]"]+)\])?\}/.source,
  liveReference: /([A-Za-z_][\w.]*)\[([^[\]"]+)\]/.source,
} as const;

export const BraceFormula = {
  compileForCell({ configs, tableKey, formula }: BraceFormulaProps): string {
    return replaceBraces({ tableKey, formula }, (brace) => {
      const reference = liveReferenceOf(configs, brace, formula);
      if (brace.isWholeColumn) return reference;
      return `SINGLE(${reference})`;
    });
  },
  // Takes a layout rather than reading the configs, so a rule can compile against a planned layout.
  compileForRule({ tableKey, formula, layout }: RuleFormulaProps): string {
    return replaceBraces({ tableKey, formula }, (brace) => {
      if (brace.tableKey !== tableKey) {
        throw new Error(
          braceMessage(
            brace,
            "names another Table, which a rule's formula can't reach",
            formula,
          ),
        );
      }
      const offset = layout.columnNamesLeftToRight.indexOf(brace.columnName);
      if (offset === -1) throw new Error(noColumnMessage(brace, formula));
      const colIndex = SheetIndex.col(layout.headerTopLeft.colIndex + offset);
      const firstBodyRow = SheetIndex.row(layout.headerTopLeft.rowIndex + 1);
      if (brace.isWholeColumn) {
        return schema.openColumnA1(colIndex, firstBodyRow);
      }
      return schema.anchoredA1(colIndex, firstBodyRow);
    });
  },
  columnsNamed({ tableKey, formula }: FormulaText): ConfiguredColumn[] {
    const columns: ConfiguredColumn[] = [];
    replaceBraces({ tableKey, formula }, (brace) => {
      columns.push({
        tableKey: brace.tableKey,
        columnName: brace.columnName,
      });
      return brace.text;
    });
    return columns;
  },
  // A reference to no configured column stays as Sheets wrote it.
  decompile({ configs, tableKey, formula }: BraceFormulaProps): string {
    const liveTables = liveTablesByName(configs);
    const { stringLiteral, liveReference } = formulaPatterns;
    const pattern = new RegExp(
      `(${stringLiteral})|SINGLE\\(\\s*${liveReference}\\s*\\)|${liveReference}`,
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
          if (column.tableKey === tableKey) return `\${${column.columnName}}`;
          return `SINGLE(${wholeColumnBrace(column)})`;
        }
        const column = configuredColumn(liveTables, table, header);
        if (column === undefined) return match;
        return wholeColumnBrace(column);
      },
    );
  },
};

// A string literal is matched first and kept, so a brace inside one stays text.
function replaceBraces(
  { tableKey, formula }: FormulaText,
  replace: (brace: Brace) => string,
): string {
  const pattern = new RegExp(
    `(${formulaPatterns.stringLiteral})|${formulaPatterns.brace}`,
    "g",
  );
  return formula.replace(
    pattern,
    (
      text: string,
      literal: string | undefined,
      first: string,
      second: string | undefined,
    ) => {
      if (literal !== undefined) return literal;
      if (second === undefined) {
        return replace({
          text,
          tableKey,
          columnName: first,
          isWholeColumn: false,
        });
      }
      return replace({
        text,
        tableKey: first,
        columnName: second,
        isWholeColumn: true,
      });
    },
  );
}

function liveReferenceOf(
  configs: BraceConfigs,
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

function braceMessage(brace: Brace, problem: string, formula: string): string {
  return `The brace ${brace.text} ${problem}, in formula ${formula}.`;
}

function noColumnMessage(brace: Brace, formula: string): string {
  return braceMessage(
    brace,
    `names no column of Table ${brace.tableKey}`,
    formula,
  );
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

function wholeColumnBrace({ tableKey, columnName }: ConfiguredColumn): string {
  return `\${${tableKey}[${columnName}]}`;
}
