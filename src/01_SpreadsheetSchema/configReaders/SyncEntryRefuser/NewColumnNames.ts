import { columnNameOfHeader } from "../columnConfigsTypes";
import { entriesOf, type WideTableEntry } from "../SyncPlanner/wideAlterations";
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

// The columns a Table entry's `newColumns` create, each by the name the regen will give its header, and the names its renames give.
export class NewColumnNames {
  private headersByColumn: Map<string, NewHeader[]>;
  private oldNamesByNewName: Map<string, string>;
  constructor({ newColumns, columns }: WideTableEntry) {
    this.headersByColumn = (newColumns ?? []).reduce(
      (byColumn, { header }, index) => {
        if (header === undefined) return byColumn;
        const column = columnNameOfHeader(header);
        const newHeader = { header, path: `newColumns[${index}].header` };
        byColumn.set(column, [...(byColumn.get(column) ?? []), newHeader]);
        return byColumn;
      },
      new Map<string, NewHeader[]>(),
    );
    this.oldNamesByNewName = entriesOf(columns).reduce(
      (byNewName, [oldName, { header }]) => {
        if (header === undefined) return byNewName;
        const newName = columnNameOfHeader(header);
        if (newName !== oldName) byNewName.set(newName, oldName);
        return byNewName;
      },
      new Map<string, string>(),
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
  // Alterations key a column by its old name for the whole run (ADR 0035), so a rename's new name names nothing yet.
  renamedRefusals({ table, path, column }: ColumnNamedAt): SyncRefusal[] {
    const oldName = this.oldNamesByNewName.get(column);
    if (oldName === undefined) return [];
    return [{ kind: "renamedColumn", table, path, column, oldName }];
  }
}
