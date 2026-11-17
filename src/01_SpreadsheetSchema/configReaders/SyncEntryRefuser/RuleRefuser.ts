import {
  rulesOf,
  type WideRule,
  type WideTableEntry,
} from "../SyncPlanner/wideAlterations";
import type { SyncRefusal } from "../syncRunResult";
import { BraceRefuser } from "./BraceRefuser";
import { type ColumnNamedAt, NewColumnNames } from "./NewColumnNames";

export interface RuleRefuserProps {
  table: string;
  entry: WideTableEntry;
  columnNamesByTable: Map<string, string[]>;
}

// Refuses a Table entry's rule naming a column neither mapped nor created by its `newColumns`, or one two new headers name.
export class RuleRefuser {
  private table: string;
  private entry: WideTableEntry;
  private newColumns: NewColumnNames;
  private columnNames: string[];
  private braceRefuser: BraceRefuser;
  constructor({ table, entry, columnNamesByTable }: RuleRefuserProps) {
    this.table = table;
    this.entry = entry;
    this.newColumns = new NewColumnNames(entry);
    this.columnNames = [
      ...new Set([
        ...(columnNamesByTable.get(table) ?? []),
        ...this.newColumns.names,
      ]),
    ];
    this.braceRefuser = new BraceRefuser(
      new Map(columnNamesByTable).set(table, this.columnNames),
      new Map([[table, this.newColumns]]),
    );
  }
  get refusals(): SyncRefusal[] {
    const { table } = this;
    return rulesOf(this.entry).flatMap(({ path, rule }) => [
      ...ruleColumnNames(rule).flatMap((column) =>
        this._columnRefusals({ table, path: `${path}.columns`, column }),
      ),
      ...ruleFormulas(rule).flatMap((formula, index) =>
        this.braceRefuser.formulaRefusals({
          table,
          tableKey: table,
          path: `${path}.condition.values[${index}]`,
          formula,
        }),
      ),
    ]);
  }
  private _columnRefusals(at: ColumnNamedAt): SyncRefusal[] {
    const ambiguous = this.newColumns.ambiguousRefusals(at);
    if (ambiguous.length > 0) return ambiguous;
    if (this.columnNames.includes(at.column)) return [];
    return [{ kind: "unknownColumn", ...at, columns: this.columnNames }];
  }
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
