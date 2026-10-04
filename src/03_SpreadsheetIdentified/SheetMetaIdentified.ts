import type { UniformRowName } from "../00_Source/CellValues/cellValues";
import { SheetMetaRaw } from "../02_SpreadsheetRaw/SheetMetaRaw";
import { TableCommonIdentified } from "./ClassBases/TableCommonIdentified";
import { ColumnMetaIdentified } from "./ColumnMetaIdentified";
import { TableIdentified } from "./TableIdentified";
import { UniformRowIdentified } from "./UniformRowIdentified";

export interface GatherDataPrerequisitesProps {
  skipFetchingProperties?: boolean;
  includeProgrammaticFacts?: boolean;
}

export class SheetMetaIdentified extends TableCommonIdentified {
  get raw(): SheetMetaRaw {
    return new SheetMetaRaw(this.tableIdentifiedProps);
  }
  get primary(): TableIdentified {
    return new TableIdentified(this.tableIdentifiedProps);
  }
  column(columnId: string): ColumnMetaIdentified {
    return new ColumnMetaIdentified({
      ...this.tableIdentifiedProps,
      columnId,
    });
  }
  isActiveColumnId(columnId: string): boolean {
    return this.raw.isActiveColumnId(columnId);
  }
  columnIdByIndex(colIndex: number): string {
    return this.raw.columnIdAt(colIndex);
  }
  uniformRow<UN extends UniformRowName>(rowName: UN): UniformRowIdentified<UN> {
    return new UniformRowIdentified({
      ...this.tableIdentifiedProps,
      uniformRowName: rowName,
    });
  }
  uniformRowByIndex(rowIndex: number): UniformRowIdentified {
    return this.uniformRow(this.schema.uniformRowNameByIndex(rowIndex));
  }
  isTableColIndex(colIndex: number): boolean {
    return this.raw.isTableColIndex(colIndex);
  }
  ensureColumnIdsAreFetched(): this {
    this._gatherDataPrerequisites();
    this.raw.ss.fetchAllGathered();
    return this;
  }
  // The columnId row sits above the table, so its filter alone returns no table metadata.
  _gatherDataPrerequisites({
    skipFetchingProperties,
  }: GatherDataPrerequisitesProps = {}): void {
    if (!skipFetchingProperties && !this.raw.primary.hasFetchedProperties) {
      this.raw.primary.gatherFetchProperties();
    }
    // Skip if a prior full-row fetch on the columnId row already covers this row.
    if (
      !this.raw.hasFetchedColumnIds &&
      !this.raw.primary.hasQueuedFullRowFetch(this.schema.colIdRowIndex)
    ) {
      this.raw.gatherFetchColumnIdsInit();
    }
  }
  gatherFetchDataPrepped(): void {
    // This is so that table dimensions and columnIndexes can be guaranteed
    // before their fetch requests are generated.
    this.fetchTargets.forEach((target) => {
      if (target.kind === "fullRow") {
        this.raw.primary.rowCommon(target.row).gatherFetchFull();
      } else if (target.kind === "fullDataColumn") {
        this.column(target.column).raw.primary.gatherFetchFull();
      } else if (target.kind === "singleCell") {
        const colIndex = this.column(target.column).colIndex;
        this.raw.primary
          .rowCommon(target.row)
          .cell(colIndex)
          .gatherFetchRange();
      } else {
        const exhaustive: never = target;
        throw new Error(`Unknown fetch target: ${JSON.stringify(exhaustive)}`);
      }
    });
  }
  addMissingColumnIds(): number {
    return this.raw.addMissingColumnIds(this.schema.idPrefix);
  }
}
