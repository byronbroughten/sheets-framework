import { columnNameOfHeader } from "../columnConfigsTypes";
import type { WideTableEntry } from "../SyncPlanner/wideAlterations";
import type { SyncRefusal } from "../syncRunResult";

export interface ColumnNamedAt {
  table: string;
  path: string;
  column: string;
}

// The columns a Table entry's `newColumns` create, each by the name the regen will give its header.
export class NewColumnNames {
  private headersByColumn: Map<string, string[]>;
  constructor({ newColumns }: WideTableEntry) {
    this.headersByColumn = (newColumns ?? []).reduce((byColumn, { header }) => {
      if (header === undefined) return byColumn;
      const column = columnNameOfHeader(header);
      byColumn.set(column, [...(byColumn.get(column) ?? []), header]);
      return byColumn;
    }, new Map<string, string[]>());
  }
  get names(): string[] {
    return [...this.headersByColumn.keys()];
  }
  // Two headers giving one name make a rule naming it ambiguous.
  ambiguousRefusals({ table, path, column }: ColumnNamedAt): SyncRefusal[] {
    const headers = this.headersByColumn.get(column) ?? [];
    if (headers.length < 2) return [];
    return [{ kind: "ambiguousColumn", table, path, column, headers }];
  }
}
