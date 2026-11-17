import type { ColumnTypeBaseline } from "../columnBaseline";
import { BraceFormula, type FormulaRenames } from "./BraceFormula";
import type { TableNow } from "./namesNow";
import { RuleRecords } from "./RuleRecords";
import {
  entriesOf,
  type WideColumnEntry,
  type WideRule,
  type WideTableEntry,
} from "./SyncPlanner/wideAlterations";
import type { SyncRewrite } from "./syncRunResult";

export interface Rename {
  written: string;
  // Undefined when the regen gives the ID no name: the entry keeps the one written.
  now: string | undefined;
  rewrite: (now: string) => SyncRewrite;
}

export interface SyncFormulaRewriterProps {
  renames: FormulaRenames;
  // The Table whose columns bare braces name; undefined for a Table this run creates, whose own columns no regen has renamed.
  tableKey: string | undefined;
  table: TableNow | undefined;
  // The name an entry takes, recording the rewrite in the run's order.
  renamed: (rename: Rename) => string;
}

export type FormulaTexts = Pick<WideColumnEntry, "formula" | "columnType">;

type Rules = NonNullable<WideTableEntry["conditionalFormats"]>;

interface RuleColumnsRewrite {
  columns: WideRule["columns"];
  at: string;
}

// Rewrites the names inside one Table's pending formula text, rule columns and rule handles.
export class SyncFormulaRewriter {
  private readonly renames: FormulaRenames;
  private readonly tableKey: string | undefined;
  private readonly table: TableNow | undefined;
  private readonly renamed: (rename: Rename) => string;
  constructor({ renames, tableKey, table, renamed }: SyncFormulaRewriterProps) {
    this.renames = renames;
    this.tableKey = tableKey;
    this.table = table;
    this.renamed = renamed;
  }
  formulaTexts(entry: FormulaTexts, at: string): FormulaTexts {
    const { formula, columnType } = entry;
    return {
      ...(formula === undefined
        ? {}
        : { formula: this._formula(formula, `${at}.formula`) }),
      ...(columnType === undefined
        ? {}
        : { columnType: this._columnType(columnType, `${at}.columnType`) }),
    };
  }
  rules(rules: Rules, at: string): Rules {
    const { add, remove, replace } = rules;
    return {
      ...rules,
      ...(add === undefined ? {} : { add: this.addedRules(add, `${at}.add`) }),
      ...(remove === undefined
        ? {}
        : {
            remove: remove.map((handle, index) =>
              this._handle(handle, `${at}.remove[${index}]`),
            ),
          }),
      ...(replace === undefined
        ? {}
        : {
            replace: entriesOf(replace).reduce<Record<string, WideRule>>(
              (rewritten, [handle, rule]) => {
                const handleNow = this._handle(handle, `${at}.replace`);
                rewritten[handleNow] = this._rule(
                  rule,
                  `${at}.replace["${handle}"]`,
                );
                return rewritten;
              },
              {},
            ),
          }),
    };
  }
  addedRules(add: WideRule[], at: string): WideRule[] {
    return add.map((rule, index) => this._rule(rule, `${at}[${index}]`));
  }
  private _columnType(
    columnType: ColumnTypeBaseline,
    at: string,
  ): ColumnTypeBaseline {
    const { rule } = columnType;
    if (rule === null) return columnType;
    return {
      ...columnType,
      rule: {
        ...rule,
        values: this._formulaValues(rule.values, `${at}.rule.values`),
      },
    };
  }
  // Only a value starting with `=` is a formula; a list item such as `Size[cm]` stays text.
  private _formulaValues(values: string[], at: string): string[] {
    return values.map((value, index) => {
      if (!value.startsWith("=")) return value;
      return this._formula(value, `${at}[${index}]`);
    });
  }
  private _formula(formula: string, at: string): string {
    return this.renamed({
      written: formula,
      now: this._formulaNow(formula),
      rewrite: (now) => ({ from: `${at} "${formula}"`, to: `"${now}"` }),
    });
  }
  private _formulaNow(formula: string): string {
    return BraceFormula.renamed({
      tableKey: this.tableKey,
      formula,
      renames: this.renames,
    });
  }
  private _rule(rule: WideRule, at: string): WideRule {
    const { columns, condition } = rule;
    const { values } = condition;
    return {
      ...rule,
      columns: this._ruleColumns({ columns, at }),
      condition: {
        ...condition,
        ...(values === undefined
          ? {}
          : {
              values: this._formulaValues(values, `${at}.condition.values`),
            }),
      },
    };
  }
  private _ruleColumns({
    columns,
    at,
  }: RuleColumnsRewrite): WideRule["columns"] {
    if (columns === "all") return columns;
    return columns.map((columnName) =>
      this.renamed({
        written: columnName,
        now: this.table?.columns.get(columnName)?.name,
        rewrite: (now) => ({
          from: `${at}.columns "${columnName}"`,
          to: `"${now}"`,
        }),
      }),
    );
  }
  // A handle changes only where the configs would record it under a renamed column or Table.
  private _handle(handle: string, at: string): string {
    return this.renamed({
      written: handle,
      now: RuleRecords.renamedHandle({
        handle,
        columnNow: (columnName) =>
          this.table?.columns.get(columnName)?.name ?? columnName,
        formulaNow: (formula) => this._formulaNow(formula),
      }),
      rewrite: (now) => ({ from: `${at} "${handle}"`, to: `"${now}"` }),
    });
  }
}
