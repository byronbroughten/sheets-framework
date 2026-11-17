import { Val } from "@byronbroughten/utils/val";

import type { ColumnTypeBaseline } from "../columnBaseline";
import { BraceFormula, type FormulaRenames } from "./BraceFormula";
import { subjectKey } from "./didntLand";
import type { NamesNow, TableNow } from "./namesNow";
import { RuleRecords } from "./RuleRecords";
import {
  entriesOf,
  type WideAlterations,
  type WideColumnEntry,
  type WideNewTable,
  type WideRule,
  type WideTabEntry,
  type WideTableEntry,
} from "./SyncPlanner/wideAlterations";
import type { SyncRewrite, SyncSubject } from "./syncRunResult";

export interface SyncAlterationsRewriterProps {
  names: NamesNow;
  // Each column's header in the regenerated configs, by column ID.
  headers: Map<string, string>;
  renames: FormulaRenames;
  sentKeys: Set<string>;
}

interface Rename {
  written: string;
  // Undefined when the regen gives the ID no name: the entry keeps the one written.
  now: string | undefined;
  rewrite: (now: string) => SyncRewrite;
}

interface BeforeRename {
  before: string | undefined;
  // Where the entry sits, as the rewrite line names it.
  at: string;
  nowOf: (name: string) => string | undefined;
}

interface TableEntryRewrite {
  tableName: string;
  entry: WideTableEntry;
  table: TableNow | undefined;
}

interface ColumnsRewrite {
  tableName: string;
  columns: NonNullable<WideTableEntry["columns"]>;
  table: TableNow | undefined;
}

type FormulaTexts = Pick<WideColumnEntry, "formula" | "columnType">;

// Where formula text sits, and the Table whose columns its bare braces name; undefined for a Table this run creates.
interface FormulaPlace {
  at: string;
  tableKey: string | undefined;
}

interface FormulaTextsRewrite extends FormulaPlace {
  entry: FormulaTexts;
}

interface FormulaValuesRewrite extends FormulaPlace {
  values: string[];
}

interface FormulaRewrite extends FormulaPlace {
  formula: string;
}

interface RulePlace extends FormulaPlace {
  // Undefined for a Table this run creates, whose columns no regen has renamed.
  table: TableNow | undefined;
}

interface RuleRewrite extends RulePlace {
  rule: WideRule;
}

interface HandleRewrite extends RulePlace {
  handle: string;
}

interface RulesRewrite {
  tableName: string;
  rules: NonNullable<WideTableEntry["conditionalFormats"]>;
  table: TableNow | undefined;
}

// Rewrites pending alterations' names, collecting a line for each rewrite it makes.
export class SyncAlterationsRewriter {
  readonly rewrites: SyncRewrite[] = [];
  private readonly names: NamesNow;
  private readonly headers: Map<string, string>;
  private readonly renames: FormulaRenames;
  private readonly sentKeys: Set<string>;
  constructor({
    names,
    headers,
    renames,
    sentKeys,
  }: SyncAlterationsRewriterProps) {
    this.names = names;
    this.headers = headers;
    this.renames = renames;
    this.sentKeys = sentKeys;
  }
  rewritten(alterations: WideAlterations): WideAlterations {
    const { tables, tabs, newTabs, ...others } = alterations;
    return {
      ...others,
      ...(tables === undefined ? {} : { tables: this._tables(tables) }),
      ...(tabs === undefined ? {} : { tabs: this._tabs(tabs) }),
      ...(newTabs === undefined
        ? {}
        : {
            newTabs: newTabs.map((newTab, index) => {
              const at = `newTabs[${index}]`;
              return {
                ...newTab,
                ...this._tabBefore(newTab, at),
                ...this._tabNewTables(newTab, at),
                tables: this._newTables(newTab.tables, `${at}.tables`),
              };
            }),
          }),
    };
  }
  private _tables(
    tables: NonNullable<WideAlterations["tables"]>,
  ): Record<string, WideTableEntry> {
    return entriesOf(tables).reduce<Record<string, WideTableEntry>>(
      (rewritten, [tableName, entry]) => {
        const table = this.names.tables.get(tableName);
        const name = this._renamed({
          written: tableName,
          now: table?.name,
          rewrite: (now) => {
            const { tableId } = Val.assert(table, `Table ${tableName}`);
            const by = this._renamedBy({ table: tableName, path: "name" });
            return {
              from: tableName,
              to: now,
              why: `tableId ${tableId}, which ${by} renamed`,
            };
          },
        });
        rewritten[name] = {
          ...rewritten[name],
          ...this._tableEntry({ tableName, entry, table }),
        };
        return rewritten;
      },
      {},
    );
  }
  private _tableEntry({
    tableName,
    entry,
    table,
  }: TableEntryRewrite): WideTableEntry {
    const { columns, newColumns, conditionalFormats } = entry;
    const rewritten: WideTableEntry = { ...entry };
    if (columns !== undefined) {
      rewritten.columns = this._columns({ tableName, columns, table });
    }
    if (newColumns !== undefined) {
      rewritten.newColumns = newColumns.map((newColumn, index) => {
        const at = `${tableName}.newColumns[${index}]`;
        return {
          ...newColumn,
          ...this._columnBefore(newColumn, at, table),
          ...this._formulaTexts({ entry: newColumn, at, tableKey: tableName }),
        };
      });
    }
    if (conditionalFormats !== undefined) {
      rewritten.conditionalFormats = this._rules({
        tableName,
        rules: conditionalFormats,
        table,
      });
    }
    return rewritten;
  }
  private _columns({
    tableName,
    columns,
    table,
  }: ColumnsRewrite): Record<string, WideColumnEntry> {
    const tableNow = table?.name ?? tableName;
    return entriesOf(columns).reduce<Record<string, WideColumnEntry>>(
      (byName, [columnName, columnEntry]) => {
        const column = table?.columns.get(columnName);
        const at = `${tableName}.columns.${columnName}`;
        const name = this._renamed({
          written: columnName,
          now: column?.name,
          rewrite: (now) => {
            const columnId = Val.assert(column, `column ${at}`).columnId;
            const header = Val.assert(
              this.headers.get(columnId),
              `header of ${columnId}`,
            );
            const by = this._renamedBy({
              table: tableName,
              column: columnName,
              path: "header",
            });
            return {
              from: at,
              to: `${tableNow}.columns.${now}`,
              why: `${columnId}, which ${by} retitled "${header}"`,
            };
          },
        });
        byName[name] = {
          ...byName[name],
          ...columnEntry,
          ...this._columnBefore(columnEntry, at, table),
          ...this._formulaTexts({
            entry: columnEntry,
            at,
            tableKey: tableName,
          }),
        };
        return byName;
      },
      {},
    );
  }
  private _columnBefore(
    entry: WideColumnEntry,
    at: string,
    table: TableNow | undefined,
  ): Pick<WideColumnEntry, "before"> {
    return this._before({
      before: entry.before,
      at,
      nowOf: (name) => table?.columns.get(name)?.name,
    });
  }
  private _formulaTexts({
    entry,
    at,
    tableKey,
  }: FormulaTextsRewrite): FormulaTexts {
    const { formula, columnType } = entry;
    return {
      ...(formula === undefined
        ? {}
        : {
            formula: this._formula({
              formula,
              at: `${at}.formula`,
              tableKey,
            }),
          }),
      ...(columnType === undefined
        ? {}
        : {
            columnType: this._columnType(columnType, {
              at: `${at}.columnType`,
              tableKey,
            }),
          }),
    };
  }
  private _columnType(
    columnType: ColumnTypeBaseline,
    { at, tableKey }: FormulaPlace,
  ): ColumnTypeBaseline {
    const { rule } = columnType;
    if (rule === null) return columnType;
    return {
      ...columnType,
      rule: {
        ...rule,
        values: this._formulaValues({
          values: rule.values,
          at: `${at}.rule.values`,
          tableKey,
        }),
      },
    };
  }
  // Only a value starting with `=` is a formula; a list item such as `Size[cm]` stays text.
  private _formulaValues({
    values,
    at,
    tableKey,
  }: FormulaValuesRewrite): string[] {
    return values.map((value, index) => {
      if (!value.startsWith("=")) return value;
      return this._formula({ formula: value, at: `${at}[${index}]`, tableKey });
    });
  }
  private _formula({ formula, at, tableKey }: FormulaRewrite): string {
    return this._renamed({
      written: formula,
      now: this._formulaNow(formula, tableKey),
      rewrite: (now) => ({ from: `${at} "${formula}"`, to: `"${now}"` }),
    });
  }
  private _formulaNow(formula: string, tableKey: string | undefined): string {
    return BraceFormula.renamed({ tableKey, formula, renames: this.renames });
  }
  private _rules({
    tableName,
    rules,
    table,
  }: RulesRewrite): NonNullable<WideTableEntry["conditionalFormats"]> {
    const { add, remove, replace } = rules;
    const at = `${tableName}.conditionalFormats`;
    const place = { tableKey: tableName, table };
    return {
      ...rules,
      ...(add === undefined
        ? {}
        : {
            add: add.map((rule, index) =>
              this._rule({ rule, at: `${at}.add[${index}]`, ...place }),
            ),
          }),
      ...(remove === undefined
        ? {}
        : {
            remove: remove.map((handle, index) =>
              this._handle({ handle, at: `${at}.remove[${index}]`, ...place }),
            ),
          }),
      ...(replace === undefined
        ? {}
        : {
            replace: entriesOf(replace).reduce<Record<string, WideRule>>(
              (rewritten, [handle, rule]) => {
                const handleNow = this._handle({
                  handle,
                  at: `${at}.replace`,
                  ...place,
                });
                rewritten[handleNow] = this._rule({
                  rule,
                  at: `${at}.replace["${handle}"]`,
                  ...place,
                });
                return rewritten;
              },
              {},
            ),
          }),
    };
  }
  private _rule({ rule, at, tableKey, table }: RuleRewrite): WideRule {
    const { columns, condition } = rule;
    const { values } = condition;
    return {
      ...rule,
      columns: this._ruleColumns(columns, at, table),
      condition: {
        ...condition,
        ...(values === undefined
          ? {}
          : {
              values: this._formulaValues({
                values,
                at: `${at}.condition.values`,
                tableKey,
              }),
            }),
      },
    };
  }
  private _ruleColumns(
    columns: WideRule["columns"],
    at: string,
    table: TableNow | undefined,
  ): WideRule["columns"] {
    if (columns === "all") return columns;
    return columns.map((columnName) =>
      this._renamed({
        written: columnName,
        now: table?.columns.get(columnName)?.name,
        rewrite: (now) => ({
          from: `${at}.columns "${columnName}"`,
          to: `"${now}"`,
        }),
      }),
    );
  }
  // A handle changes only where the configs would record it under a renamed column or Table.
  private _handle({ handle, at, tableKey, table }: HandleRewrite): string {
    return this._renamed({
      written: handle,
      now: RuleRecords.renamedHandle({
        handle,
        columnNow: (columnName) =>
          table?.columns.get(columnName)?.name ?? columnName,
        formulaNow: (formula) => this._formulaNow(formula, tableKey),
      }),
      rewrite: (now) => ({ from: `${at} "${handle}"`, to: `"${now}"` }),
    });
  }
  private _tabs(
    tabs: NonNullable<WideAlterations["tabs"]>,
  ): Record<string, WideTabEntry> {
    return entriesOf(tabs).reduce<Record<string, WideTabEntry>>(
      (rewritten, [tabName, entry]) => {
        const tab = this.names.tabs.get(tabName);
        const name = this._renamed({
          written: tabName,
          now: tab?.name,
          rewrite: (now) => {
            const { sheetGid, title } = Val.assert(tab, `tab ${tabName}`);
            const by = this._renamedBy({ tab: tabName, path: "title" });
            return {
              from: `tabs.${tabName}`,
              to: `tabs.${now}`,
              why: `gid ${sheetGid}, which ${by} retitled "${title}"`,
            };
          },
        });
        const at = `tabs.${tabName}`;
        rewritten[name] = {
          ...rewritten[name],
          ...entry,
          ...this._tabBefore(entry, at),
          ...this._tabNewTables(entry, at),
        };
        return rewritten;
      },
      {},
    );
  }
  private _tabBefore(
    entry: WideTabEntry,
    at: string,
  ): Pick<WideTabEntry, "before"> {
    return this._before({
      before: entry.before,
      at,
      nowOf: (name) => this.names.tabs.get(name)?.name,
    });
  }
  private _tabNewTables(
    entry: WideTabEntry,
    at: string,
  ): Pick<WideTabEntry, "newTables"> {
    const { newTables } = entry;
    if (newTables === undefined) return {};
    return { newTables: this._newTables(newTables, `${at}.newTables`) };
  }
  // A new Table's own columns have no previous name, so only what its text names elsewhere is rewritten.
  private _newTables(newTables: WideNewTable[], at: string): WideNewTable[] {
    return newTables.map((newTable, tableIndex) => {
      const tableAt = `${at}[${tableIndex}]`;
      const { columns, conditionalFormats } = newTable;
      const add = conditionalFormats?.add;
      const place = { tableKey: undefined, table: undefined };
      return {
        ...newTable,
        columns: columns.map((column, index) => ({
          ...column,
          ...this._formulaTexts({
            entry: column,
            at: `${tableAt}.columns[${index}]`,
            tableKey: undefined,
          }),
        })),
        ...(add === undefined
          ? {}
          : {
              conditionalFormats: {
                ...conditionalFormats,
                add: add.map((rule, index) =>
                  this._rule({
                    rule,
                    at: `${tableAt}.conditionalFormats.add[${index}]`,
                    ...place,
                  }),
                ),
              },
            }),
      };
    });
  }
  private _before({ before, at, nowOf }: BeforeRename): { before?: string } {
    if (before === undefined) return {};
    return {
      before: this._renamed({
        written: before,
        now: nowOf(before),
        rewrite: (now) => ({
          from: `${at}.before "${before}"`,
          to: `"${now}"`,
        }),
      }),
    };
  }
  // The name an entry takes, recording the rewrite when the regen renamed it.
  private _renamed({ written, now, rewrite }: Rename): string {
    if (now === undefined || now === written) return written;
    this.rewrites.push(rewrite(now));
    return now;
  }
  private _renamedBy(subject: SyncSubject): string {
    return this.sentKeys.has(subjectKey(subject)) ? "this run" : "a person";
  }
}
