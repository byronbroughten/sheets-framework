import { SheetMetaRaw } from "../02_SpreadsheetRaw/SheetMetaRaw";
import { TableCommonIdentified } from "./ClassBases/TableCommonIdentified";
import { TableIdentified } from "./TableIdentified";

export class SheetMetaIdentified extends TableCommonIdentified {
  get raw(): SheetMetaRaw {
    return new SheetMetaRaw(this.tableIdentifiedProps);
  }
  get primary(): TableIdentified {
    return new TableIdentified(this.tableIdentifiedProps);
  }
  ensureColumnIdsAreFetched(): this {
    this.primary.columnResolver.gatherDataPrerequisites();
    this.raw.ss.fetchAllGathered();
    return this;
  }
  gatherFetchDataPrepped(): void {
    // This is so that table dimensions and columnIndexes can be guaranteed
    // before their fetch requests are generated.
    this.fetchTargets.forEach((target) => {
      if (target.kind === "fullRow") {
        this.raw.primary.rowCommon(target.row).gatherFetchFull();
      } else if (target.kind === "fullDataColumn") {
        this.primary.column(target.column).raw.gatherFetchFull();
      } else if (target.kind === "singleCell") {
        const colIndex = this.primary.column(target.column).colIndex;
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
