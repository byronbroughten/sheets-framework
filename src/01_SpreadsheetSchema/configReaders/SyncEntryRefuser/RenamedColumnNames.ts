import { columnNameOfHeader } from "../columnConfigsTypes";
import {
  entriesOf,
  type WideAlterations,
  type WideTableEntry,
} from "../SyncPlanner/wideAlterations";
import type { SyncRefusal } from "../syncRunResult";
import type { ColumnNamedAt } from "./namedColumnRefusals";

// The names a Table entry's renames give its columns, each to the column's old name.
export class RenamedColumnNames {
  private oldNamesByNewName: Map<string, string>;
  constructor({ columns }: WideTableEntry) {
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
  static byTable({ tables }: WideAlterations): Map<string, RenamedColumnNames> {
    return new Map(
      entriesOf(tables).map(([table, entry]) => [
        table,
        new RenamedColumnNames(entry),
      ]),
    );
  }
  // Alterations key a column by its old name for the whole run (ADR 0035), so a rename's new name names nothing yet.
  refusals({ table, path, column }: ColumnNamedAt): SyncRefusal[] {
    const oldName = this.oldNamesByNewName.get(column);
    if (oldName === undefined) return [];
    return [{ kind: "renamedColumn", table, path, column, oldName }];
  }
}
