import { Val } from "@byronbroughten/utils/val";

import type { FormulaRenames } from "./BraceFormula";
import { subjectKey } from "./didntLand";
import type { NamesNow, TableNow } from "./namesNow";
import { type Rename, SyncFormulaRewriter } from "./SyncFormulaRewriter";
import {
  entriesOf,
  type WideAlterations,
  type WideColumnEntry,
  type WideNewTable,
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
  formulas: SyncFormulaRewriter;
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
    const formulas = this._formulaRewriter(tableName, table);
    if (columns !== undefined) {
      rewritten.columns = this._columns({
        tableName,
        columns,
        table,
        formulas,
      });
    }
    if (newColumns !== undefined) {
      rewritten.newColumns = newColumns.map((newColumn, index) => {
        const at = `${tableName}.newColumns[${index}]`;
        return {
          ...newColumn,
          ...this._columnBefore(newColumn, at, table),
          ...formulas.formulaTexts(newColumn, at),
        };
      });
    }
    if (conditionalFormats !== undefined) {
      rewritten.conditionalFormats = formulas.rules(
        conditionalFormats,
        `${tableName}.conditionalFormats`,
      );
    }
    return rewritten;
  }
  private _columns({
    tableName,
    columns,
    table,
    formulas,
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
          ...formulas.formulaTexts(columnEntry, at),
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
    const formulas = this._formulaRewriter(undefined, undefined);
    return newTables.map((newTable, tableIndex) => {
      const tableAt = `${at}[${tableIndex}]`;
      const { columns, conditionalFormats } = newTable;
      const add = conditionalFormats?.add;
      return {
        ...newTable,
        columns: columns.map((column, index) => ({
          ...column,
          ...formulas.formulaTexts(column, `${tableAt}.columns[${index}]`),
        })),
        ...(add === undefined
          ? {}
          : {
              conditionalFormats: {
                ...conditionalFormats,
                add: formulas.addedRules(
                  add,
                  `${tableAt}.conditionalFormats.add`,
                ),
              },
            }),
      };
    });
  }
  private _formulaRewriter(
    tableKey: string | undefined,
    table: TableNow | undefined,
  ): SyncFormulaRewriter {
    return new SyncFormulaRewriter({
      renames: this.renames,
      tableKey,
      table,
      renamed: (rename) => this._renamed(rename),
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
