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
export interface RuleDecompileProps extends RuleFormulaProps {
  isOneColumn: boolean;
}

interface RuleReference {
  sheetPrefix: string | undefined;
  start: CellAddress;
  end: CellAddress | undefined;
}

interface CellAddress {
  isColAbsolute: boolean;
  colLetter: string;
  isRowAbsolute: boolean;
  rowNumber: number | undefined;
}

interface RuleBraceContext {
  tableKey: string;
  columnNamesByLetter: Map<string, string>;
  firstBodyRowNumber: number;
  isOneColumn: boolean;
}

interface Brace {
  text: string;
  tableKey: string;
  columnName: string;
  isWholeColumn: boolean;
}

export interface RuleLayout {
  headerTopLeft: { rowIndex: SheetRowIndex; colIndex: SheetColIndex };
  // Undefined at a column with no ID yet, which a run that mints none leaves out of the configs.
  columnNamesLeftToRight: readonly (string | undefined)[];
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
  // The lookarounds keep a function name such as LOG10( from reading as a cell.
  a1Reference:
    /(?<![\w$.'])((?:'(?:[^']|'')*'|[A-Za-z_][\w.]*)!)?(\$?)([A-Z]{1,3})(\$?)(\d+)?(?::(\$?)([A-Z]{1,3})(\$?)(\d+)?)?(?![\w$(])/
      .source,
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
  // The inverse of compileForRule; undefined when a reference has no brace form.
  decompileForRule({
    tableKey,
    formula,
    layout,
    isOneColumn,
  }: RuleDecompileProps): string | undefined {
    const columnNamesByLetter = layout.columnNamesLeftToRight.reduce(
      (byLetter, columnName, offset) => {
        if (columnName === undefined) return byLetter;
        const colIndex = SheetIndex.col(layout.headerTopLeft.colIndex + offset);
        byLetter.set(schema.columnLetter(colIndex), columnName);
        return byLetter;
      },
      new Map<string, string>(),
    );
    const firstBodyRowNumber = layout.headerTopLeft.rowIndex + 2;
    const pattern = new RegExp(
      `(${formulaPatterns.stringLiteral})|${formulaPatterns.a1Reference}`,
      "g",
    );
    let isExpressible = true;
    const decompiled = formula.replace(
      pattern,
      (
        match: string,
        literal: string | undefined,
        sheetPrefix: string | undefined,
        ...addressParts: unknown[]
      ) => {
        if (literal !== undefined) return match;
        const reference = ruleReferenceOf(sheetPrefix, addressParts);
        if (reference === undefined) return match;
        const brace = ruleBraceOf(reference, {
          tableKey,
          columnNamesByLetter,
          firstBodyRowNumber,
          isOneColumn,
        });
        if (brace === undefined) isExpressible = false;
        return brace ?? match;
      },
    );
    if (!isExpressible) return undefined;
    return decompiled;
  },
  columnsNamed({ tableKey, formula }: FormulaText): ConfiguredColumn[] {
    return [...formula.matchAll(bracePattern())].flatMap(
      ([text, literal, first, second]) => {
        const brace = braceOf(tableKey, { text, literal, first, second });
        if (brace === undefined) return [];
        return [{ tableKey: brace.tableKey, columnName: brace.columnName }];
      },
    );
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

function replaceBraces(
  { tableKey, formula }: FormulaText,
  replace: (brace: Brace) => string,
): string {
  return formula.replace(
    bracePattern(),
    (
      text: string,
      literal: string | undefined,
      first: string | undefined,
      second: string | undefined,
    ) => {
      const brace = braceOf(tableKey, { text, literal, first, second });
      if (brace === undefined) return text;
      return replace(brace);
    },
  );
}

// A string literal is matched first and kept, so a brace inside one stays text.
function bracePattern(): RegExp {
  return new RegExp(
    `(${formulaPatterns.stringLiteral})|${formulaPatterns.brace}`,
    "g",
  );
}

interface BraceMatch {
  text: string;
  literal: string | undefined;
  first: string | undefined;
  second: string | undefined;
}

// Undefined for a string literal, which isn't a brace.
function braceOf(
  tableKey: string,
  { text, literal, first, second }: BraceMatch,
): Brace | undefined {
  if (literal !== undefined || first === undefined) return undefined;
  if (second === undefined) {
    return { text, tableKey, columnName: first, isWholeColumn: false };
  }
  return { text, tableKey: first, columnName: second, isWholeColumn: true };
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

// Undefined for a bare word such as TRUE, which has neither a row nor a colon.
function ruleReferenceOf(
  sheetPrefix: string | undefined,
  addressParts: unknown[],
): RuleReference | undefined {
  const start = cellAddressOf(addressParts.slice(0, 4));
  if (start === undefined) return undefined;
  const end = cellAddressOf(addressParts.slice(4, 8));
  if (end === undefined && start.rowNumber === undefined) return undefined;
  return { sheetPrefix, start, end };
}

// Takes one address's four groups: column $, column letters, row $, row digits.
function cellAddressOf(groups: unknown[]): CellAddress | undefined {
  const [colAbs, colLetter, rowAbs, rowDigits] = groups;
  if (typeof colLetter !== "string") return undefined;
  let rowNumber: number | undefined;
  if (typeof rowDigits === "string") rowNumber = Number(rowDigits);
  return {
    isColAbsolute: colAbs === "$",
    colLetter,
    isRowAbsolute: rowAbs === "$",
    rowNumber,
  };
}

function ruleBraceOf(
  { sheetPrefix, start, end }: RuleReference,
  context: RuleBraceContext,
): string | undefined {
  if (sheetPrefix !== undefined) return undefined;
  const columnName = context.columnNamesByLetter.get(start.colLetter);
  if (columnName === undefined) return undefined;
  if (end === undefined) {
    if (start.isRowAbsolute) return undefined;
    if (start.rowNumber !== context.firstBodyRowNumber) return undefined;
    if (!start.isColAbsolute && !context.isOneColumn) return undefined;
    return `\${${columnName}}`;
  }
  const isOpenColumn =
    start.isColAbsolute &&
    start.isRowAbsolute &&
    start.rowNumber === context.firstBodyRowNumber &&
    end.isColAbsolute &&
    end.colLetter === start.colLetter &&
    end.rowNumber === undefined;
  if (!isOpenColumn) return undefined;
  return wholeColumnBrace({ tableKey: context.tableKey, columnName });
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
