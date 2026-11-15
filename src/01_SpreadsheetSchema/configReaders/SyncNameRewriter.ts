import { Val } from "@byronbroughten/utils/val";

import { subjectKey } from "./didntLand";
import type { IdNameMapGeneric } from "./idNameMapTypes";
import type { ColumnConfigsGeneric } from "./makeConfigs";
import {
  entriesOf,
  type WideAlterations,
  type WideColumnEntry,
  type WideTabEntry,
  type WideTableEntry,
} from "./SyncPlanner/wideAlterations";
import type {
  SyncMovedName,
  SyncRename,
  SyncRewrite,
  SyncSubject,
} from "./syncRunResult";

export interface SyncNameRewriterProps {
  // The alterations still pending after this run's landed ones left.
  alterations: WideAlterations;
  idNameMaps: { previous: IdNameMapGeneric; regenerated: IdNameMapGeneric };
  columnConfigs: {
    previous: ColumnConfigsGeneric;
    regenerated: ColumnConfigsGeneric;
  };
  sent: SyncSubject[];
  landed: SyncSubject[];
}

interface ColumnNow {
  columnId: string;
  name: string;
}

interface TableNow {
  tableId: string;
  name: string;
  // Keyed by each column's previous name.
  columns: Map<string, ColumnNow>;
}

interface TabNow {
  sheetGid: string;
  name: string;
  title: string;
}

// Each Table and tab both maps hold, keyed by its previous name, with the name the regen gives it now.
interface NamesNow {
  tables: Map<string, TableNow>;
  tabs: Map<string, TabNow>;
}

interface RewriteContext {
  names: NamesNow;
  headers: Map<string, string>;
  sentKeys: Set<string>;
  rewrites: SyncRewrite[];
}

/**
 * Rewrites the names in pending alterations that a regen renamed, through
 * the ID each named in the previous ID-name map (ADR 0035): Table, column
 * and tab keys and every `before:`. Formula text and rule handles are left
 * as written. It also words the Configs section's renames, the names that
 * now point at another column, and each rewrite, and shows a pending
 * entry's subject under the name it now has, with the name it was written.
 * Endpoint and chore code are never rewritten: tsc fails there and retarget fixes it.
 */
export class SyncNameRewriter {
  readonly alterations: WideAlterations;
  readonly rewrites: SyncRewrite[];
  readonly renamed: SyncRename[];
  readonly namesMoved: SyncMovedName[];
  private readonly names: NamesNow;
  private readonly landedKeys: Set<string>;
  constructor({
    alterations,
    idNameMaps,
    columnConfigs,
    sent,
    landed,
  }: SyncNameRewriterProps) {
    this.names = namesNow(idNameMaps);
    this.landedKeys = new Set(landed.map(subjectKey));
    const sentKeys = new Set(sent.map(subjectKey));
    const regeneratedHeaders = headersByColumnId(columnConfigs.regenerated);
    const context: RewriteContext = {
      names: this.names,
      headers: regeneratedHeaders,
      sentKeys,
      rewrites: [],
    };
    this.alterations = rewrittenAlterations(alterations, context);
    this.rewrites = context.rewrites;
    this.renamed = personRenames({
      names: this.names,
      headers: {
        previous: headersByColumnId(columnConfigs.previous),
        regenerated: regeneratedHeaders,
      },
      sentKeys,
    });
    this.namesMoved = movedNames(idNameMaps);
  }
  static init(props: SyncNameRewriterProps): SyncNameRewriter {
    return new SyncNameRewriter(props);
  }
  // A landed entry left the file rather than being rewritten, so its subject keeps the name it was sent under.
  asWritten(subject: SyncSubject): SyncSubject {
    if ("tab" in subject || subject.column === undefined) return subject;
    if (this.landedKeys.has(subjectKey(subject))) return subject;
    const table = this.names.tables.get(subject.table);
    const column = table?.columns.get(subject.column);
    if (table === undefined || column === undefined) return subject;
    if (column.name === subject.column) return subject;
    const { path } = subject;
    return {
      table: table.name,
      column: column.name,
      rewrittenFrom: { wrote: subject.column, columnId: column.columnId },
      ...(path === undefined ? {} : { path }),
    };
  }
}

function namesNow({
  previous,
  regenerated,
}: SyncNameRewriterProps["idNameMaps"]): NamesNow {
  const tables = new Map<string, TableNow>();
  for (const [tableId, previousTable] of Object.entries(previous.tables)) {
    const table = regenerated.tables[tableId];
    if (table === undefined) continue;
    const namesById = new Map(
      Object.entries(table.columns ?? {}).map(([name, columnId]) => [
        columnId,
        name,
      ]),
    );
    const columns = new Map<string, ColumnNow>();
    for (const [previousName, columnId] of Object.entries(
      previousTable.columns ?? {},
    )) {
      const name = namesById.get(columnId);
      if (name !== undefined) columns.set(previousName, { columnId, name });
    }
    tables.set(previousTable.name, { tableId, name: table.name, columns });
  }
  const tabs = new Map<string, TabNow>();
  for (const [sheetGid, previousTab] of Object.entries(previous.tabs)) {
    const tab = regenerated.tabs[sheetGid];
    if (tab === undefined) continue;
    tabs.set(previousTab.name, { sheetGid, name: tab.name, title: tab.title });
  }
  return { tables, tabs };
}

function headersByColumnId(
  columnConfigs: ColumnConfigsGeneric,
): Map<string, string> {
  return new Map(
    Object.values(columnConfigs).flatMap((tableColumns) =>
      Object.values(tableColumns).map(
        ({ columnId, header }): [string, string] => [columnId, header],
      ),
    ),
  );
}

function renamedBy(isSent: boolean): string {
  return isSent ? "this run" : "a person";
}

function rewrittenAlterations(
  alterations: WideAlterations,
  context: RewriteContext,
): WideAlterations {
  const { tables, tabs, newTabs, ...others } = alterations;
  return {
    ...others,
    ...(tables === undefined
      ? {}
      : { tables: rewrittenTables(tables, context) }),
    ...(tabs === undefined ? {} : { tabs: rewrittenTabs(tabs, context) }),
    ...(newTabs === undefined
      ? {}
      : {
          newTabs: newTabs.map((newTab, index) => ({
            ...newTab,
            ...rewrittenTabBefore(newTab, `newTabs[${index}]`, context),
          })),
        }),
  };
}

function rewrittenTables(
  tables: NonNullable<WideAlterations["tables"]>,
  context: RewriteContext,
): Record<string, WideTableEntry> {
  return entriesOf(tables).reduce<Record<string, WideTableEntry>>(
    (rewritten, [tableName, entry]) => {
      const table = context.names.tables.get(tableName);
      const name = table?.name ?? tableName;
      if (table !== undefined && name !== tableName) {
        const isSent = context.sentKeys.has(
          subjectKey({ table: tableName, path: "name" }),
        );
        context.rewrites.push({
          from: tableName,
          to: name,
          why: `tableId ${table.tableId}, which ${renamedBy(isSent)} renamed`,
        });
      }
      rewritten[name] = {
        ...rewritten[name],
        ...rewrittenTableEntry({ tableName, entry, table, context }),
      };
      return rewritten;
    },
    {},
  );
}

interface TableEntryRewrite {
  tableName: string;
  entry: WideTableEntry;
  table: TableNow | undefined;
  context: RewriteContext;
}

function rewrittenTableEntry({
  tableName,
  entry,
  table,
  context,
}: TableEntryRewrite): WideTableEntry {
  const { columns, newColumns } = entry;
  const newName = table?.name ?? tableName;
  const rewritten: WideTableEntry = { ...entry };
  if (columns !== undefined) {
    rewritten.columns = entriesOf(columns).reduce<
      Record<string, WideColumnEntry>
    >((byName, [columnName, columnEntry]) => {
      const column = table?.columns.get(columnName);
      const name = column?.name ?? columnName;
      if (column !== undefined && name !== columnName) {
        const isSent = context.sentKeys.has(
          subjectKey({ table: tableName, column: columnName, path: "header" }),
        );
        const header = Val.assert(
          context.headers.get(column.columnId),
          `header of ${column.columnId}`,
        );
        context.rewrites.push({
          from: `${tableName}.columns.${columnName}`,
          to: `${newName}.columns.${name}`,
          why: `${column.columnId}, which ${renamedBy(isSent)} retitled "${header}"`,
        });
      }
      byName[name] = {
        ...byName[name],
        ...columnEntry,
        ...rewrittenColumnBefore({
          entry: columnEntry,
          at: `${tableName}.columns.${columnName}`,
          table,
          rewrites: context.rewrites,
        }),
      };
      return byName;
    }, {});
  }
  if (newColumns !== undefined) {
    rewritten.newColumns = newColumns.map((newColumn, index) => ({
      ...newColumn,
      ...rewrittenColumnBefore({
        entry: newColumn,
        at: `${tableName}.newColumns[${index}]`,
        table,
        rewrites: context.rewrites,
      }),
    }));
  }
  return rewritten;
}

interface ColumnBeforeRewrite {
  entry: WideColumnEntry;
  // Where the entry sits, as the rewrite line names it.
  at: string;
  table: TableNow | undefined;
  rewrites: SyncRewrite[];
}

function rewrittenColumnBefore({
  entry,
  at,
  table,
  rewrites,
}: ColumnBeforeRewrite): Pick<WideColumnEntry, "before"> {
  const { before } = entry;
  if (before === undefined) return {};
  const name = table?.columns.get(before)?.name ?? before;
  if (name !== before) {
    rewrites.push({ from: `${at}.before "${before}"`, to: `"${name}"` });
  }
  return { before: name };
}

function rewrittenTabs(
  tabs: NonNullable<WideAlterations["tabs"]>,
  context: RewriteContext,
): Record<string, WideTabEntry> {
  return entriesOf(tabs).reduce<Record<string, WideTabEntry>>(
    (rewritten, [tabName, entry]) => {
      const tab = context.names.tabs.get(tabName);
      const name = tab?.name ?? tabName;
      if (tab !== undefined && name !== tabName) {
        const isSent = context.sentKeys.has(
          subjectKey({ tab: tabName, path: "title" }),
        );
        context.rewrites.push({
          from: `tabs.${tabName}`,
          to: `tabs.${name}`,
          why: `gid ${tab.sheetGid}, which ${renamedBy(isSent)} retitled "${tab.title}"`,
        });
      }
      rewritten[name] = {
        ...rewritten[name],
        ...entry,
        ...rewrittenTabBefore(entry, `tabs.${tabName}`, context),
      };
      return rewritten;
    },
    {},
  );
}

function rewrittenTabBefore(
  entry: WideTabEntry,
  at: string,
  context: RewriteContext,
): Pick<WideTabEntry, "before"> {
  const { before } = entry;
  if (before === undefined) return {};
  const name = context.names.tabs.get(before)?.name ?? before;
  if (name !== before) {
    context.rewrites.push({
      from: `${at}.before "${before}"`,
      to: `"${name}"`,
    });
  }
  return { before: name };
}

interface PersonRenamesProps {
  names: NamesNow;
  headers: { previous: Map<string, string>; regenerated: Map<string, string> };
  sentKeys: Set<string>;
}

// A run's own header send is in its plan already, so only a person's rename is listed here.
function personRenames({
  names,
  headers,
  sentKeys,
}: PersonRenamesProps): SyncRename[] {
  return [...names.tables].flatMap(([tableName, table]) =>
    [...table.columns].flatMap(([from, { columnId, name }]): SyncRename[] => {
      if (name === from) return [];
      const header = subjectKey({
        table: tableName,
        column: from,
        path: "header",
      });
      if (sentKeys.has(header)) return [];
      return [
        {
          table: table.name,
          from,
          to: name,
          headerFrom: Val.assert(
            headers.previous.get(columnId),
            `previous header of ${columnId}`,
          ),
          headerTo: Val.assert(
            headers.regenerated.get(columnId),
            `header of ${columnId}`,
          ),
        },
      ];
    }),
  );
}

// Every column name both maps hold for one Table that now names another column ID.
function movedNames({
  previous,
  regenerated,
}: SyncNameRewriterProps["idNameMaps"]): SyncMovedName[] {
  return Object.entries(previous.tables).flatMap(([tableId, previousTable]) => {
    const table = regenerated.tables[tableId];
    if (table === undefined) return [];
    return Object.entries(previousTable.columns ?? {}).flatMap(
      ([column, wasId]): SyncMovedName[] => {
        const nowId = table.columns?.[column];
        if (nowId === undefined || nowId === wasId) return [];
        return [{ table: table.name, column, wasId, nowId }];
      },
    );
  });
}
