import { columnNameOfHeader } from "../columnConfigsTypes";
import {
  newColumnPath,
  type WideTableEntry,
} from "../SyncPlanner/wideAlterations";
import type { SyncRefusal } from "../syncRunResult";

export interface ColumnNamedAt {
  table: string;
  path: string;
  column: string;
}

interface NewHeader {
  header: string;
  // The header's path in the Table entry, as `newColumns[0].header`.
  path: string;
}

// The columns a Table entry's `newColumns` create, each by the name the regen will give its header.
export class NewColumnNames {
  private headersByColumn: Map<string, NewHeader[]>;
  constructor({ newColumns }: WideTableEntry) {
    this.headersByColumn = (newColumns ?? []).reduce(
      (byColumn, { header }, index) => {
        if (header === undefined) return byColumn;
        const column = columnNameOfHeader(header);
        const newHeader = { header, path: `${newColumnPath(index)}.header` };
        byColumn.set(column, [...(byColumn.get(column) ?? []), newHeader]);
        return byColumn;
      },
      new Map<string, NewHeader[]>(),
    );
  }
  get names(): string[] {
    return [...this.headersByColumn.keys()];
  }
  // The refetch names each column by its header, so two headers giving one name can't both be found.
  repeatRefusals(table: string): SyncRefusal[] {
    return [...this.headersByColumn].flatMap(
      ([column, [first, ...repeats]]) => {
        if (first === undefined) return [];
        return repeats.map(({ path }): SyncRefusal => ({
          kind: "repeatedNewColumn",
          table,
          path,
          firstPath: first.path,
          column,
        }));
      },
    );
  }
  // Two headers giving one name make a rule naming it ambiguous.
  ambiguousRefusals({ table, path, column }: ColumnNamedAt): SyncRefusal[] {
    const headers = (this.headersByColumn.get(column) ?? []).map(
      ({ header }) => header,
    );
    if (headers.length < 2) return [];
    return [{ kind: "ambiguousColumn", table, path, column, headers }];
  }
}
