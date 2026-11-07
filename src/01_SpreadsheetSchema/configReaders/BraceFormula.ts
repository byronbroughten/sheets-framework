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

// tableName is the codebase name of the Table whose formula this is.
interface FormulaText {
  tableName: string;
  formula: string;
}
interface BraceFormulaProps extends FormulaText {
  configs: BraceConfigs;
}
interface RuleFormulaProps extends FormulaText {
  layout: RuleLayout;
}

// `${columnName}` is the formula's own row; `${tableName[columnName]}` the whole column.
interface Brace {
  text: string;
  tableName: string;
  columnName: string;
  isQualified: boolean;
}

// The Table's top-left cell is its header row's first cell; columns are codebase names, left to right.
export interface RuleLayout {
  topLeft: { rowIndex: SheetRowIndex; colIndex: SheetColIndex };
  columnNames: readonly string[];
}

interface LiveTable {
  tableName: string;
  columnNamesByHeader: Map<string, string>;
}

const stringLiteral = /"(?:[^"]|"")*"/.source;
const braceText = /\$\{([^{}[\]"]+?)(?:\[([^{}[\]"]+)\])?\}/.source;
const liveReference = /([A-Za-z_][\w.]*)\[([^[\]"]+)\]/.source;

export const BraceFormula = {
  compileForCell({ configs, tableName, formula }: BraceFormulaProps): string {
    return replaceBraces({ tableName, formula }, (brace) => {
      const reference = liveReferenceOf(configs, brace, formula);
      if (brace.isQualified) return reference;
      return `SINGLE(${reference})`;
    });
  },
  // Takes a layout rather than reading the configs, so a rule can compile against a planned layout.
  compileForRule({ tableName, formula, layout }: RuleFormulaProps): string {
    return replaceBraces({ tableName, formula }, (brace) => {
      if (brace.tableName !== tableName) {
        throw new Error(
          `The brace ${brace.text} names another Table, which a rule's formula can't reach, in formula ${formula}.`,
        );
      }
      const offset = layout.columnNames.indexOf(brace.columnName);
      if (offset === -1) throw new Error(noColumnMessage(brace, formula));
      const colIndex = SheetIndex.col(layout.topLeft.colIndex + offset);
      const firstBodyRow = SheetIndex.row(layout.topLeft.rowIndex + 1);
      if (!brace.isQualified) return schema.anchoredA1(colIndex, firstBodyRow);
      const letter = schema.columnLetter(colIndex);
      return `$${letter}$${firstBodyRow + 1}:$${letter}`;
    });
  },
  // A reference to no configured column stays as Sheets wrote it.
  decompile({ configs, tableName, formula }: BraceFormulaProps): string {
    const liveTables = liveTablesByName(configs);
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
          const liveTable = liveTables.get(singleTable);
          const columnName = liveTable?.columnNamesByHeader.get(singleHeader);
          if (liveTable === undefined || columnName === undefined) return match;
          if (liveTable.tableName === tableName) return `\${${columnName}}`;
          return `SINGLE(\${${liveTable.tableName}[${columnName}]})`;
        }
        const liveTable = liveTables.get(table);
        const columnName = liveTable?.columnNamesByHeader.get(header);
        if (liveTable === undefined || columnName === undefined) return match;
        return `\${${liveTable.tableName}[${columnName}]}`;
      },
    );
  },
};

// A string literal is matched first and kept, so a brace inside one stays text.
function replaceBraces(
  { tableName, formula }: FormulaText,
  replace: (brace: Brace) => string,
): string {
  const pattern = new RegExp(`(${stringLiteral})|${braceText}`, "g");
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
          tableName,
          columnName: first,
          isQualified: false,
        });
      }
      return replace({
        text,
        tableName: first,
        columnName: second,
        isQualified: true,
      });
    },
  );
}

function liveReferenceOf(
  configs: BraceConfigs,
  brace: Brace,
  formula: string,
): string {
  const tableConfig = configs.tableConfigs[brace.tableName];
  if (tableConfig === undefined) {
    throw new Error(
      `The brace ${brace.text} names no Table, in formula ${formula}.`,
    );
  }
  const columnConfig =
    configs.columnConfigs[brace.tableName]?.[brace.columnName];
  if (columnConfig === undefined) {
    throw new Error(noColumnMessage(brace, formula));
  }
  return `${tableConfig.tableName}[${columnConfig.header}]`;
}

function noColumnMessage(brace: Brace, formula: string): string {
  return `The brace ${brace.text} names no column of Table ${brace.tableName}, in formula ${formula}.`;
}

function liveTablesByName(configs: BraceConfigs): Map<string, LiveTable> {
  return Object.entries(configs.tableConfigs).reduce(
    (byName, [tableName, tableConfig]) => {
      const columnNamesByHeader = Object.entries(
        configs.columnConfigs[tableName] ?? {},
      ).reduce((byHeader, [columnName, columnConfig]) => {
        byHeader.set(columnConfig.header, columnName);
        return byHeader;
      }, new Map<string, string>());
      byName.set(tableConfig.tableName, { tableName, columnNamesByHeader });
      return byName;
    },
    new Map<string, LiveTable>(),
  );
}
