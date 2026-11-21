import { columnNameOfHeader } from "../columnConfigsTypes.ts";
import {
  newColumnsOf,
  type WideNewTable,
  type WideTableEntry,
} from "../SyncPlanner/wideAlterations.ts";
import type { SyncRefusal } from "../syncRunResult.ts";

interface NewHeader {
  header: string | undefined;
  // The header's path in its alterations, as `newColumns[0].columns[1].header`.
  path: string;
}

interface NewTableHeadersProps {
  newTable: WideNewTable;
  // Where the new Table sits in its alterations, as `newTables[0]`.
  at: string;
}

// The columns a Table entry's `newColumns` or a new Table's columns create, each by the name the regen will give its header.
export class NewColumnNames {
  private pathsByColumn: Map<string, string[]>;
  constructor(headers: NewHeader[]) {
    this.pathsByColumn = headers.reduce((byColumn, { header, path }) => {
      if (header === undefined) return byColumn;
      const column = columnNameOfHeader(header);
      byColumn.set(column, [...(byColumn.get(column) ?? []), path]);
      return byColumn;
    }, new Map<string, string[]>());
  }
  static ofTableEntry(entry: WideTableEntry): NewColumnNames {
    return new NewColumnNames(
      newColumnsOf(entry).map(({ newColumn: { header }, path }) => ({
        header,
        path: `${path}.header`,
      })),
    );
  }
  static ofNewTable({ newTable, at }: NewTableHeadersProps): NewColumnNames {
    return new NewColumnNames(
      newTable.columns.map(({ header }, index) => ({
        header,
        path: `${at}.columns[${index}].header`,
      })),
    );
  }
  get names(): string[] {
    return [...this.pathsByColumn.keys()];
  }
  // The refetch names each column by its header, so two headers giving one name can't both be found.
  repeatRefusals(table: string): SyncRefusal[] {
    return [...this.pathsByColumn].flatMap(
      ([column, [firstPath, ...repeats]]) => {
        if (firstPath === undefined) return [];
        return repeats.map((path): SyncRefusal => ({
          kind: "repeatedNewColumn",
          table,
          path,
          firstPath,
          column,
        }));
      },
    );
  }
}
