import {
  ruleRemovalLines,
  ruleReplacementLines,
  unplacedNewTable,
} from "../SyncPlanner";
import {
  entriesOf,
  rulesOf,
  type WideNewTable,
  type WideRule,
  type WideTableEntry,
} from "../SyncPlanner/wideAlterations";
import type { SyncRefusal } from "../syncRunResult";
import { BraceRefuser } from "./BraceRefuser";
import { namedColumnRefusals } from "./namedColumnRefusals";
import { type ColumnNamedAt, NewColumnNames } from "./NewColumnNames";
import type { RenamedColumnNames } from "./RenamedColumnNames";

export interface RuleRefuserProps {
  table: string;
  entry: WideTableEntry;
  columnNamesByTable: Map<string, string[]>;
  // Every Table entry's renames, since a brace may name another Table's column.
  renamedColumnsByTable: Map<string, RenamedColumnNames>;
}

export interface NewTableRuleRefuserProps {
  tab: string;
  newTable: WideNewTable;
  // Where the new Table sits in its alterations, as `newTables[0]`.
  at: string;
  columnNamesByTable: Map<string, string[]>;
  renamedColumnsByTable: Map<string, RenamedColumnNames>;
}

interface RuleRefuserConstructorProps extends RuleRefuserProps {
  // What a refusal names: the Table, or a new Table's tab.
  refusedAs: string;
  // Where the entry sits in its alterations, ending in a dot, or "" for a top-level entry.
  pathPrefix: string;
}

// Refuses a Table entry's rule naming a column neither mapped nor created by its `newColumns`, one two new headers name, or one by its same-run new name; and a handle both removed and replaced.
export class RuleRefuser {
  private refusedAs: string;
  private tableKey: string;
  private pathPrefix: string;
  private entry: WideTableEntry;
  private newColumns: NewColumnNames;
  private renamedColumns: RenamedColumnNames | undefined;
  private columnNames: string[];
  private braceRefuser: BraceRefuser;
  private constructor({
    table,
    refusedAs,
    pathPrefix,
    entry,
    columnNamesByTable,
    renamedColumnsByTable,
  }: RuleRefuserConstructorProps) {
    this.refusedAs = refusedAs;
    this.tableKey = table;
    this.pathPrefix = pathPrefix;
    this.entry = entry;
    this.newColumns = new NewColumnNames(entry);
    this.renamedColumns = renamedColumnsByTable.get(table);
    this.columnNames = [
      ...new Set([
        ...(columnNamesByTable.get(table) ?? []),
        ...this.newColumns.names,
      ]),
    ];
    this.braceRefuser = new BraceRefuser({
      columnNamesByTable: new Map(columnNamesByTable).set(
        table,
        this.columnNames,
      ),
      newColumnsByTable: new Map([[table, this.newColumns]]),
      renamedColumnsByTable,
    });
  }
  static ofTableEntry(props: RuleRefuserProps): RuleRefuser {
    return new RuleRefuser({
      ...props,
      refusedAs: props.table,
      pathPrefix: "",
    });
  }
  // A new Table has only the columns its headers give, so a live Table its name matches lends it none, nor its renames.
  static ofNewTable({
    tab,
    newTable,
    at,
    columnNamesByTable,
    renamedColumnsByTable,
  }: NewTableRuleRefuserProps): RuleRefuser {
    const { table } = unplacedNewTable(newTable);
    const withoutLive = new Map(columnNamesByTable);
    withoutLive.delete(table);
    const withoutLiveRenames = new Map(renamedColumnsByTable);
    withoutLiveRenames.delete(table);
    return new RuleRefuser({
      table,
      refusedAs: tab,
      pathPrefix: `${at}.`,
      entry: {
        newColumns: newTable.columns.map(({ header }) => ({ header })),
        conditionalFormats: newTable.conditionalFormats,
      },
      columnNamesByTable: withoutLive,
      renamedColumnsByTable: withoutLiveRenames,
    });
  }
  get refusals(): SyncRefusal[] {
    const { refusedAs: table, tableKey, pathPrefix } = this;
    const ruleRefusals = rulesOf(this.entry).flatMap(
      ({ path: rulePath, rule }) => {
        const path = `${pathPrefix}${rulePath}`;
        return [
          ...ruleColumnNames(rule).flatMap((column) =>
            this._columnRefusals({ table, path: `${path}.columns`, column }),
          ),
          ...ruleFormulas(rule).flatMap((formula, index) =>
            this.braceRefuser.formulaRefusals({
              table,
              tableKey,
              path: `${path}.condition.values[${index}]`,
              formula,
            }),
          ),
        ];
      },
    );
    return [...this._removedAndReplacedRefusals, ...ruleRefusals];
  }
  private _columnRefusals(at: ColumnNamedAt): SyncRefusal[] {
    return namedColumnRefusals(
      at,
      {
        names: this.columnNames,
        newColumns: this.newColumns,
        renamedColumns: this.renamedColumns,
      },
      () => ({ kind: "unknownColumn", ...at, columns: this.columnNames }),
    );
  }
  // The update's index counts only the removals below it, so a handle in both would overwrite its neighbour.
  private get _removedAndReplacedRefusals(): SyncRefusal[] {
    const { remove = [], replace } = this.entry.conditionalFormats ?? {};
    return entriesOf(replace).flatMap(([handle]): SyncRefusal[] => {
      if (!remove.includes(handle)) return [];
      return [
        {
          kind: "removedAndReplaced",
          table: this.refusedAs,
          path: `${this.pathPrefix}${ruleReplacementLines.path({ handle })}`,
          removedAt: `${this.pathPrefix}${ruleRemovalLines.path({ handle })}`,
        },
      ];
    });
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
