import { BraceFormula } from "../BraceFormula";
import { unplacedNewTable } from "../SyncPlanner";
import {
  entriesOf,
  type WideNewTable,
  type WideTableEntry,
} from "../SyncPlanner/wideAlterations";
import type { SyncRefusal } from "../syncRunResult";
import type { NewColumnNames } from "./NewColumnNames";

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

// Refuses a formula's malformed braces, and the ones naming a column it doesn't know or can't tell apart.
export class BraceRefuser {
  private columnNamesByTable: Map<string, string[]>;
  private newColumnsByTable: Map<string, NewColumnNames>;
  constructor(
    columnNamesByTable: Map<string, string[]>,
    newColumnsByTable: Map<string, NewColumnNames>,
  ) {
    this.columnNamesByTable = columnNamesByTable;
    this.newColumnsByTable = newColumnsByTable;
  }
  tableEntryRefusals(table: string, entry: WideTableEntry): SyncRefusal[] {
    return entriesOf(entry.columns).flatMap(([column, { formula }]) => {
      if (formula === undefined) return [];
      return this.formulaRefusals({
        table,
        tableKey: table,
        path: `${column}.formula`,
        formula,
      });
    });
  }
  // A brace naming the new Table resolves against its planned columns, each named as the regen will name its header.
  newTableRefusals({ tab, newTable, at }: NewTableBraceProps): SyncRefusal[] {
    const { table: tableKey, columns } = unplacedNewTable(newTable);
    const withPlanned = new BraceRefuser(
      new Map(this.columnNamesByTable).set(
        tableKey,
        columns.map(({ column }) => column),
      ),
      this.newColumnsByTable,
    );
    return newTable.columns.flatMap(({ formula }, columnIndex) => {
      if (formula === undefined) return [];
      return withPlanned.formulaRefusals({
        table: tab,
        tableKey,
        path: `${at}.columns[${columnIndex}].formula`,
        formula,
      });
    });
  }
  formulaRefusals({
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
      const ambiguous =
        this.newColumnsByTable
          .get(tableKey)
          ?.ambiguousRefusals({ table, path, column: columnName }) ?? [];
      if (ambiguous.length > 0) return ambiguous;
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
