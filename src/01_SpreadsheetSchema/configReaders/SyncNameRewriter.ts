import { Val } from "@byronbroughten/utils/val";

import type { FormulaRenames } from "./BraceFormula";
import type { ConfigSnapshots } from "./configSnapshots";
import { subjectKey } from "./didntLand";
import type { ColumnConfigsGeneric, TableConfigsBase } from "./makeConfigs";
import { movedNames, type NamesNow, namesNow } from "./namesNow";
import { SyncAlterationsRewriter } from "./SyncAlterationsRewriter";
import type { WideAlterations } from "./SyncPlanner/wideAlterations";
import type {
  SyncMovedName,
  SyncRename,
  SyncRewrite,
  SyncSubject,
} from "./syncRunResult";

export interface SyncNameRewriterProps {
  // The alterations still pending after this run's landed ones left.
  alterations: WideAlterations;
  configs: ConfigSnapshots;
  sent: SyncSubject[];
  landed: SyncSubject[];
}

// Each ID's name in the previous configs and in the regenerated ones.
interface NamesAcrossRegen {
  previous: Map<string, string>;
  regenerated: Map<string, string>;
}

interface PersonRenamesProps {
  names: NamesNow;
  headers: NamesAcrossRegen;
  sentKeys: Set<string>;
}

interface FormulaRenamesProps {
  names: NamesNow;
  // By column ID.
  headers: NamesAcrossRegen;
  // Each Table's name on the sheet, by Table ID.
  tableNames: NamesAcrossRegen;
}

// Rewrites pending alterations' names through the IDs the previous ID-name map gave them (ADR 0035), and words the report's renames.
export class SyncNameRewriter {
  readonly alterations: WideAlterations;
  readonly rewrites: SyncRewrite[];
  readonly renamed: SyncRename[];
  readonly namesMoved: SyncMovedName[];
  private readonly names: NamesNow;
  private readonly landedKeys: Set<string>;
  constructor({ alterations, configs, sent, landed }: SyncNameRewriterProps) {
    const { previous, regenerated } = configs;
    const idNameMaps = {
      previous: previous.idNameMap,
      regenerated: regenerated.idNameMap,
    };
    this.names = namesNow(idNameMaps);
    this.landedKeys = new Set(landed.map(subjectKey));
    const sentKeys = new Set(sent.map(subjectKey));
    const headers = {
      previous: headersByColumnId(previous.columnConfigs),
      regenerated: headersByColumnId(regenerated.columnConfigs),
    };
    const rewriter = new SyncAlterationsRewriter({
      names: this.names,
      headers: headers.regenerated,
      renames: formulaRenames({
        names: this.names,
        headers,
        tableNames: {
          previous: tableNamesById(previous.tableConfigs),
          regenerated: tableNamesById(regenerated.tableConfigs),
        },
      }),
      sentKeys,
    });
    this.alterations = rewriter.rewritten(alterations);
    this.rewrites = rewriter.rewrites;
    this.renamed = personRenames({ names: this.names, headers, sentKeys });
    this.namesMoved = movedNames(idNameMaps);
  }
  static init(props: SyncNameRewriterProps): SyncNameRewriter {
    return new SyncNameRewriter(props);
  }
  // A landed entry left the file rather than being rewritten, so its subject keeps the name it was sent under.
  asRewritten(subject: SyncSubject): SyncSubject {
    if (this.landedKeys.has(subjectKey(subject))) return subject;
    if ("tab" in subject) {
      const tab = this.names.tabs.get(subject.tab);
      if (tab === undefined) return subject;
      return { ...subject, tab: tab.name };
    }
    const table = this.names.tables.get(subject.table);
    if (table === undefined) return subject;
    const renamedTable = { ...subject, table: table.name };
    if (subject.column === undefined) return renamedTable;
    const column = table.columns.get(subject.column);
    if (column === undefined || column.name === subject.column) {
      return renamedTable;
    }
    const { path } = subject;
    return {
      table: table.name,
      column: column.name,
      rewrittenFrom: { wrote: subject.column, columnId: column.columnId },
      ...(path === undefined ? {} : { path }),
    };
  }
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

function tableNamesById(tableConfigs: TableConfigsBase): Map<string, string> {
  return new Map(
    Object.values(tableConfigs).map(({ tableId, tableName }) => [
      tableId,
      tableName,
    ]),
  );
}

// A Table the previous configs held without API access has no live name to rewrite.
function formulaRenames({
  names,
  headers,
  tableNames,
}: FormulaRenamesProps): FormulaRenames {
  return [...names.tables].reduce<FormulaRenames>(
    (renames, [previousName, { tableId, name, columns }]) => {
      renames.names.set(previousName, {
        name,
        columns: new Map(
          [...columns].map(([from, column]) => [from, column.name]),
        ),
      });
      const liveFrom = tableNames.previous.get(tableId);
      const liveTo = tableNames.regenerated.get(tableId);
      if (liveFrom === undefined || liveTo === undefined) return renames;
      renames.liveNames.set(liveFrom, {
        name: liveTo,
        columns: [...columns.values()].reduce((byHeader, { columnId }) => {
          const from = headers.previous.get(columnId);
          const to = headers.regenerated.get(columnId);
          if (from !== undefined && to !== undefined) byHeader.set(from, to);
          return byHeader;
        }, new Map<string, string>()),
      });
      return renames;
    },
    { names: new Map(), liveNames: new Map() },
  );
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
