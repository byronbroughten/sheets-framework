import { BraceFormula } from "../BraceFormula";
import { columnNameOfHeader } from "../columnConfigsTypes";
import { unplacedNewTable } from "../SyncPlanner";
import {
  entriesOf,
  rulesOf,
  type WideNewTable,
  type WideRule,
  type WideTableEntry,
} from "../SyncPlanner/wideAlterations";
import type { SyncRefusal } from "../syncRunResult";

interface NewTableBraceProps {
  tab: string;
  newTable: WideNewTable;
  at: string;
}

interface FormulaBraceProps {
  // The name a refusal gives; `tableKey` is the one a bare brace resolves against.
  table: string;
  tableKey: string;
  path: string;
  formula: string;
}

interface RuleFormulaProps {
  table: string;
  path: string;
  formula: string;
  headersByColumn: Map<string, string[]>;
}

// Refuses a formula's malformed braces, and the braces and rule columns naming a column it doesn't know or can't tell apart.
export class BraceRefuser {
  private columnNamesByTable: Map<string, string[]>;
  constructor(columnNamesByTable: Map<string, string[]>) {
    this.columnNamesByTable = columnNamesByTable;
  }
  tableEntryRefusals(table: string, entry: WideTableEntry): SyncRefusal[] {
    const formulaRefusals = entriesOf(entry.columns).flatMap(
      ([column, { formula }]) => {
        if (formula === undefined) return [];
        return this._formulaRefusals({
          table,
          tableKey: table,
          path: `${column}.formula`,
          formula,
        });
      },
    );
    return [...formulaRefusals, ...this._ruleRefusals(table, entry)];
  }
  // A rule can name a column the entry's `newColumns` create, by the name the regen will give its header.
  private _ruleRefusals(table: string, entry: WideTableEntry): SyncRefusal[] {
    const headersByColumn = newHeadersByColumn(entry);
    const columnNames = [
      ...new Set([
        ...(this.columnNamesByTable.get(table) ?? []),
        ...headersByColumn.keys(),
      ]),
    ];
    const withNew = new BraceRefuser(
      new Map(this.columnNamesByTable).set(table, columnNames),
    );
    return rulesOf(entry).flatMap(({ path: rulePath, rule }) => {
      const columnRefusals = ruleColumnNames(rule).flatMap((column) =>
        columnNameRefusals({
          table,
          path: `${rulePath}.columns`,
          column,
          columnNames,
          headersByColumn,
        }),
      );
      const braceRefusals = ruleFormulas(rule).flatMap((formula, index) =>
        withNew._ruleFormulaRefusals({
          table,
          path: `${rulePath}.condition.values[${index}]`,
          formula,
          headersByColumn,
        }),
      );
      return [...columnRefusals, ...braceRefusals];
    });
  }
  private _ruleFormulaRefusals({
    table,
    path,
    formula,
    headersByColumn,
  }: RuleFormulaProps): SyncRefusal[] {
    const ambiguous = BraceFormula.columnsNamed({
      tableKey: table,
      formula,
    }).flatMap(({ tableKey, columnName }) => {
      if (tableKey !== table) return [];
      return ambiguousRefusals({
        table,
        path,
        column: columnName,
        headersByColumn,
      });
    });
    return [
      ...this._formulaRefusals({ table, tableKey: table, path, formula }),
      ...ambiguous,
    ];
  }
  // A brace naming the new Table resolves against its planned columns, each named as the regen will name its header.
  newTableRefusals({ tab, newTable, at }: NewTableBraceProps): SyncRefusal[] {
    const { table: tableKey, columns } = unplacedNewTable(newTable);
    const withPlanned = new BraceRefuser(
      new Map(this.columnNamesByTable).set(
        tableKey,
        columns.map(({ column }) => column),
      ),
    );
    return newTable.columns.flatMap(({ formula }, columnIndex) => {
      if (formula === undefined) return [];
      return withPlanned._formulaRefusals({
        table: tab,
        tableKey,
        path: `${at}.columns[${columnIndex}].formula`,
        formula,
      });
    });
  }
  private _formulaRefusals({
    table,
    tableKey: ownTableKey,
    path,
    formula,
  }: FormulaBraceProps): SyncRefusal[] {
    const malformed = BraceFormula.malformedBraces({ formula }).map(
      (brace): SyncRefusal => ({ kind: "malformedBrace", table, path, brace }),
    );
    const unknown = BraceFormula.columnsNamed({
      tableKey: ownTableKey,
      formula,
    }).flatMap(({ brace, tableKey, columnName }): SyncRefusal[] => {
      const columnNames = this.columnNamesByTable.get(tableKey);
      if (columnNames?.includes(columnName)) return [];
      return [
        {
          kind: "unknownBrace",
          table,
          path,
          brace,
          braceTable: tableKey,
          columns: columnNames,
        },
      ];
    });
    return [...malformed, ...unknown];
  }
}

// Each new column's headers by the name the regen gives them; two headers on one name make it ambiguous.
function newHeadersByColumn({
  newColumns,
}: WideTableEntry): Map<string, string[]> {
  return (newColumns ?? []).reduce((byColumn, { header }) => {
    if (header === undefined) return byColumn;
    const column = columnNameOfHeader(header);
    byColumn.set(column, [...(byColumn.get(column) ?? []), header]);
    return byColumn;
  }, new Map<string, string[]>());
}

function ruleColumnNames({ columns }: WideRule): string[] {
  if (columns === "all") return [];
  return columns;
}

// Only a custom formula's value compiles its braces; any other value is a literal.
function ruleFormulas({ condition }: WideRule): string[] {
  if (condition.type !== "CUSTOM_FORMULA") return [];
  return condition.values ?? [];
}

interface ColumnNameCheck {
  table: string;
  path: string;
  column: string;
  headersByColumn: Map<string, string[]>;
}

interface RuleColumnCheck extends ColumnNameCheck {
  columnNames: string[];
}

function columnNameRefusals({
  columnNames,
  ...check
}: RuleColumnCheck): SyncRefusal[] {
  const ambiguous = ambiguousRefusals(check);
  if (ambiguous.length > 0) return ambiguous;
  if (columnNames.includes(check.column)) return [];
  const { table, path, column } = check;
  return [{ kind: "unknownColumn", table, path, column, columns: columnNames }];
}

function ambiguousRefusals({
  table,
  path,
  column,
  headersByColumn,
}: ColumnNameCheck): SyncRefusal[] {
  const headers = headersByColumn.get(column) ?? [];
  if (headers.length < 2) return [];
  return [{ kind: "ambiguousColumn", table, path, column, headers }];
}
