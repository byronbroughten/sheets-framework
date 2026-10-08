import { TableColumnResolverRaw } from "../../02_SpreadsheetRaw/TableRaw/TableColumnResolverRaw";
import { TableCommonIdentified } from "../ClassBases/TableCommonIdentified";

export interface GatherDataPrerequisitesProps {
  skipFetchingProperties?: boolean;
  includeProgrammaticFacts?: boolean;
}

export class TableColumnResolverIdentified extends TableCommonIdentified {
  get raw(): TableColumnResolverRaw {
    return new TableColumnResolverRaw(this.tableIdentifiedProps);
  }
  hasColumnId(columnId: string): boolean {
    return this.raw.hasColumnId(columnId);
  }
  columnIdAt(colIndex: number): string {
    return this.raw.columnIdAt(colIndex);
  }
  isTableColIndex(colIndex: number): boolean {
    return this.raw.isTableColIndex(colIndex);
  }
  // The columnId row sits above the table, so its filter alone returns no table metadata.
  gatherDataPrerequisites({
    skipFetchingProperties,
  }: GatherDataPrerequisitesProps = {}): void {
    const table = this.raw.table;
    if (!skipFetchingProperties && !table.hasFetchedProperties) {
      table.gatherFetchProperties();
    }
    // Skip if a prior full-row fetch on the columnId row already covers this row.
    if (
      !this.raw.hasFetchedColumnIds &&
      !table.hasQueuedFullRowFetch(this.schema.colIdRowIndex)
    ) {
      this.raw.gatherFetchColumnIds();
    }
  }
}
