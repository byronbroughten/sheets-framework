import { Val } from "@byronbroughten/utils/val";

import { subjectKey } from "./didntLand";
import type { NamesNow, TableNow } from "./namesNow";
import {
  entriesOf,
  type WideAlterations,
  type WideColumnEntry,
  type WideTabEntry,
  type WideTableEntry,
} from "./SyncPlanner/wideAlterations";
import type { SyncRewrite, SyncSubject } from "./syncRunResult";

export interface SyncAlterationsRewriterProps {
  names: NamesNow;
  // Each column's header in the regenerated configs, by column ID.
  headers: Map<string, string>;
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

// Rewrites pending alterations' names, collecting a line for each rewrite it makes.
export class SyncAlterationsRewriter {
  readonly rewrites: SyncRewrite[] = [];
  private readonly names: NamesNow;
  private readonly headers: Map<string, string>;
  private readonly sentKeys: Set<string>;
  constructor({ names, headers, sentKeys }: SyncAlterationsRewriterProps) {
    this.names = names;
    this.headers = headers;
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
            newTabs: newTabs.map((newTab, index) => ({
              ...newTab,
              ...this._tabBefore(newTab, `newTabs[${index}]`),
            })),
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
    const { columns, newColumns } = entry;
    const rewritten: WideTableEntry = { ...entry };
    if (columns !== undefined) {
      rewritten.columns = this._columns({ tableName, columns, table });
    }
    if (newColumns !== undefined) {
      rewritten.newColumns = newColumns.map((newColumn, index) => ({
        ...newColumn,
        ...this._columnBefore(
          newColumn,
          `${tableName}.newColumns[${index}]`,
          table,
        ),
      }));
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
        rewritten[name] = {
          ...rewritten[name],
          ...entry,
          ...this._tabBefore(entry, `tabs.${tabName}`),
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
