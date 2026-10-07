import { dimensionIds } from "../01_SpreadsheetSchema/dimensionIds";
import { TableBaseRaw } from "./ClassBases/TableBaseRaw";
import type { ColumnProfileRaw } from "./ColumnProfileRaw";
import { TableRaw } from "./TableRaw";

// Reads only, so it extends the base rather than TableCommonRaw and its queued writes.
export class TableProfileRaw extends TableBaseRaw {
  private get table(): TableRaw {
    return new TableRaw(this.tableRawProps);
  }
  get columnIds(): string[] {
    const { table } = this;
    return table.fullTableColIndexes
      .map((colIndex) => table.columnResolver.columnIdAt(colIndex))
      .filter((columnId) => columnId !== "");
  }
  idPrefix(): string | undefined {
    const columnIdsByPrefix = this._columnIdsByIdPrefix();
    if (columnIdsByPrefix.size === 0) return undefined;
    if (columnIdsByPrefix.size > 1) {
      throw new Error(
        mixedIdPrefixMessage(this.table.title, columnIdsByPrefix),
      );
    }
    return columnIdsByPrefix.keys().next().value;
  }
  columnById(columnId: string): ColumnProfileRaw {
    const { table } = this;
    return table.column(table.columnResolver.colIndexOf(columnId)).profile;
  }
  private _columnIdsByIdPrefix(): Map<string, string[]> {
    const columnIdsByPrefix = new Map<string, string[]>();
    this.columnIds.forEach((columnId) => {
      const idPrefix = dimensionIds.colIdPrefixOrUndefined(columnId);
      if (idPrefix === undefined) return;
      const columnIds = columnIdsByPrefix.get(idPrefix) ?? [];
      columnIds.push(columnId);
      columnIdsByPrefix.set(idPrefix, columnIds);
    });
    return columnIdsByPrefix;
  }
}

function mixedIdPrefixMessage(
  sheetTitle: string,
  columnIdsByPrefix: Map<string, string[]>,
): string {
  const prefixParts = [...columnIdsByPrefix.entries()].map(
    ([idPrefix, columnIds]) =>
      `"${idPrefix}" (${columnIds.length} column${
        columnIds.length === 1 ? "" : "s"
      }: ${columnIds.join(", ")})`,
  );
  return (
    `Sheet "${sheetTitle}" has column IDs with more than one ID prefix: ` +
    `${prefixParts.join("; ")}. Clear the stray column ID cells so the next ` +
    `sync can mint new ones.`
  );
}
