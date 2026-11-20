import { BraceFormula } from "../BraceFormula";
import { unplacedNewTable } from "../SyncPlanner";
import {
  entriesOf,
  type WideNewTable,
  type WideTableEntry,
} from "../SyncPlanner/wideAlterations";
import type { SyncRefusal } from "../syncRunResult";
import { namedColumnRefusals } from "./namedColumnRefusals";
import type { NewColumnNames } from "./NewColumnNames";
import type { RenamedColumnNames } from "./RenamedColumnNames";

interface BraceRefuserProps {
  columnNamesByTable: Map<string, string[]>;
  newColumnsByTable?: Map<string, NewColumnNames>;
  renamedColumnsByTable?: Map<string, RenamedColumnNames>;
}

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
  private renamedColumnsByTable: Map<string, RenamedColumnNames>;
  constructor({
    columnNamesByTable,
    newColumnsByTable = new Map(),
    renamedColumnsByTable = new Map(),
  }: BraceRefuserProps) {
    this.columnNamesByTable = columnNamesByTable;
    this.newColumnsByTable = newColumnsByTable;
    this.renamedColumnsByTable = renamedColumnsByTable;
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
    const withPlanned = new BraceRefuser({
      columnNamesByTable: new Map(this.columnNamesByTable).set(
        tableKey,
        columns.map(({ column }) => column),
      ),
      newColumnsByTable: this.newColumnsByTable,
      renamedColumnsByTable: this.renamedColumnsByTable,
    });
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
      const columnNames = this.columnNamesByTable.get(tableKey);
      return namedColumnRefusals(
        { table, path, column: columnName },
        {
          names: columnNames,
          newColumns: this.newColumnsByTable.get(tableKey),
          renamedColumns: this.renamedColumnsByTable.get(tableKey),
        },
        () => ({
          kind: "unknownBrace",
          table,
          path,
          brace,
          braceTable: tableKey,
          columns: columnNames,
        }),
      );
    });
    return [...malformed, ...unknown];
  }
}
