import { BraceFormula } from "../BraceFormula";
import type { TabsByTable } from "../SyncPlanner";
import type { SyncRefusal, SyncSubject } from "../syncRunResult";
import {
  entriesOf,
  moveBeforePath,
  moveColumnPath,
  newColumnsBeforesOf,
  newColumnsOf,
  type WideAlterations,
  type WideColumnEntry,
  type WideColumnGroupMove,
  type WideNewTable,
  type WideTableEntry,
} from "./wideAlterations";

interface DeletionRefuserProps {
  alterations: WideAlterations;
  tabsByTable: TabsByTable;
}

interface Deletion {
  target: string;
  subject: SyncSubject;
  deletedAt: string;
}

interface EnclosingTableProps {
  table: string;
  entry: WideTableEntry;
  at: string;
}

interface Touch {
  targets: string[];
  touchedBy: string;
  takenTable?: string;
}

// Refuses a deletion beside another entry touching what it deletes, naming both.
export class DeletionRefuser {
  private alterations: WideAlterations;
  private tabsByTable: TabsByTable;
  constructor(props: DeletionRefuserProps) {
    this.alterations = props.alterations;
    this.tabsByTable = props.tabsByTable;
  }
  get refusals(): SyncRefusal[] {
    const touches = this._alterationTouches;
    return this._deletions.flatMap((deletion) =>
      touches.flatMap(({ targets, touchedBy, takenTable }) => {
        if (!targets.includes(deletion.target)) return [];
        return [
          {
            kind: "touchesDeletion" as const,
            deleted: deletion.subject,
            deletedAt: deletion.deletedAt,
            touchedBy,
            ...(takenTable === undefined ? {} : { takenTable }),
          },
        ];
      }),
    );
  }
  private get _deletions(): Deletion[] {
    const tableDeletions = entriesOf(this.alterations.tables).flatMap(
      ([table, entry]) => [
        ...(entry.deleted === true
          ? [
              {
                target: tableTarget(table),
                subject: { table },
                deletedAt: `tables.${table}`,
              },
            ]
          : []),
        ...entriesOf(entry.columns).flatMap(([column, columnEntry]) => {
          if (columnEntry.deleted !== true) return [];
          return [
            {
              target: columnTarget(table, column),
              subject: { table, column },
              deletedAt: `tables.${table}.columns.${column}`,
            },
          ];
        }),
      ],
    );
    const tabDeletions = entriesOf(this.alterations.tabs).flatMap(
      ([tab, entry]) => {
        if (entry.deleted !== true) return [];
        return [
          {
            target: tabTarget(tab),
            subject: { tab },
            deletedAt: `tabs.${tab}`,
          },
        ];
      },
    );
    return [...tableDeletions, ...tabDeletions];
  }
  private get _alterationTouches(): Touch[] {
    const tableTouches = entriesOf(this.alterations.tables).flatMap(
      ([table, entry]) => this._tableEntryTouches(table, entry),
    );
    const tabTouches = entriesOf(this.alterations.tabs).flatMap(
      ([tab, entry]) => {
        if (entry.deleted === true) return [];
        const at = `tabs.${tab}`;
        return [
          { targets: [tabTarget(tab)], touchedBy: at },
          ...beforeTabTouches(entry.before, `${at}.before`),
          ...this._newTablesTouches(entry.newTables ?? [], `${at}.newTables`),
        ];
      },
    );
    const newTabTouches = (this.alterations.newTabs ?? []).flatMap(
      (newTab, index) => {
        const at = `newTabs[${index}]`;
        return [
          ...beforeTabTouches(newTab.before, `${at}.before`),
          ...this._newTablesTouches(newTab.tables, `${at}.tables`),
        ];
      },
    );
    return [...tableTouches, ...tabTouches, ...newTabTouches];
  }
  private _tableEntryTouches(table: string, entry: WideTableEntry): Touch[] {
    const at = `tables.${table}`;
    if (entry.deleted === true) return this._takenTableTouches(table, at);
    return [
      ...this._enclosingTableTouches({ table, entry, at }),
      ...entriesOf(entry.columns).flatMap(([column, columnEntry]) =>
        this._columnEntryTouches(table, column, columnEntry),
      ),
      ...newColumnsBeforesOf(entry).map(({ before, path }) => ({
        targets: this._columnTargets(table, before),
        touchedBy: `${at}.${path}`,
      })),
      ...newColumnsOf(entry).flatMap(({ newColumn, path }) =>
        this._formulaTouches(table, newColumn, `${at}.${path}`),
      ),
      ...this._moveTouches(table, entry.moves ?? []),
      ...this._ruleTouches(table, entry.conditionalFormats, at),
    ];
  }
  // A column entry's touch already reaches its Table and tab, so a Table entry of only columns adds a second refusal for the same conflict.
  private _enclosingTableTouches({
    table,
    entry,
    at,
  }: EnclosingTableProps): Touch[] {
    if (holdsOnlyColumns(entry)) return [];
    return [{ targets: this._tableTargets(table), touchedBy: at }];
  }
  private _moveTouches(table: string, moves: WideColumnGroupMove[]): Touch[] {
    const at = `tables.${table}`;
    return moves.flatMap(({ columns, before }, index) => [
      ...columns.map((column, columnIndex) => ({
        targets: this._columnTargets(table, column),
        touchedBy: `${at}.${moveColumnPath({ index, columnIndex })}`,
      })),
      {
        targets: this._columnTargets(table, before),
        touchedBy: `${at}.${moveBeforePath(index)}`,
      },
    ]);
  }
  private _columnEntryTouches(
    table: string,
    column: string,
    entry: WideColumnEntry,
  ): Touch[] {
    const at = `tables.${table}.columns.${column}`;
    if (entry.deleted === true) return this._takenTableTouches(table, at);
    return [
      { targets: this._columnTargets(table, column), touchedBy: at },
      ...this._formulaTouches(table, entry, at),
    ];
  }
  // A Table or column deletion touches only its tab, whose deletion already takes the Table.
  private _takenTableTouches(table: string, touchedBy: string): Touch[] {
    const targets = this._tabTargets(table);
    if (targets.length === 0) return [];
    return [{ targets, touchedBy, takenTable: table }];
  }
  private _formulaTouches(
    table: string,
    { formula }: WideColumnEntry,
    at: string,
  ): Touch[] {
    if (formula === undefined) return [];
    return this._braceTouches(table, formula, `${at}.formula`);
  }
  private _ruleTouches(
    table: string,
    formats: WideTableEntry["conditionalFormats"],
    tableAt: string,
  ): Touch[] {
    if (formats === undefined) return [];
    const at = `${tableAt}.conditionalFormats`;
    const added = (formats.add ?? []).map(
      (rule, index) => [`${at}.add[${index}]`, rule] as const,
    );
    const replaced = entriesOf(formats.replace).map(
      ([handle, rule]) => [`${at}.replace["${handle}"]`, rule] as const,
    );
    return [...added, ...replaced].flatMap(([ruleAt, rule]) => [
      ...(rule.columns === "all"
        ? []
        : rule.columns.map((column) => ({
            targets: this._columnTargets(table, column),
            touchedBy: `${ruleAt}.columns`,
          }))),
      ...(rule.condition.values ?? []).flatMap((value) =>
        this._braceTouches(table, value, `${ruleAt}.condition`),
      ),
    ]);
  }
  private _newTablesTouches(
    newTables: WideNewTable[],
    tablesAt: string,
  ): Touch[] {
    return newTables.flatMap((newTable, tableIndex) =>
      newTable.columns.flatMap(({ formula }, columnIndex) => {
        if (formula === undefined) return [];
        return this._braceTouches(
          newTable.name,
          formula,
          `${tablesAt}[${tableIndex}].columns[${columnIndex}].formula`,
        );
      }),
    );
  }
  private _braceTouches(
    tableKey: string,
    formula: string,
    touchedBy: string,
  ): Touch[] {
    return BraceFormula.columnsNamed({ tableKey, formula }).map(
      ({ tableKey: table, columnName }) => ({
        targets: this._columnTargets(table, columnName),
        touchedBy,
      }),
    );
  }
  private _columnTargets(table: string, column: string): string[] {
    return [columnTarget(table, column), ...this._tableTargets(table)];
  }
  private _tableTargets(table: string): string[] {
    return [tableTarget(table), ...this._tabTargets(table)];
  }
  private _tabTargets(table: string): string[] {
    const tab = this.tabsByTable[table];
    return tab === undefined ? [] : [tabTarget(tab)];
  }
}

function holdsOnlyColumns(entry: WideTableEntry): boolean {
  const { columns, ...rest } = entry;
  const hasColumns = entriesOf(columns).length > 0;
  return (
    hasColumns && Object.values(rest).every((value) => value === undefined)
  );
}

function beforeTabTouches(
  before: string | undefined,
  touchedBy: string,
): Touch[] {
  if (before === undefined) return [];
  return [{ targets: [tabTarget(before)], touchedBy }];
}

function columnTarget(table: string, column: string): string {
  return `column:${table}/${column}`;
}

function tableTarget(table: string): string {
  return `table:${table}`;
}

function tabTarget(tab: string): string {
  return `tab:${tab}`;
}
