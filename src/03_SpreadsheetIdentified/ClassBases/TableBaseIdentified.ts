import { Val } from "@byronbroughten/utils/val";

import {
  tableConfigsByTableId,
  type TableName,
} from "../../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import { TableSchema } from "../../01_SpreadsheetSchema/configReaders/TableSchema";
import { TableBaseRaw } from "../../02_SpreadsheetRaw/ClassBases/TableBaseRaw";
import { SheetRaw } from "../../02_SpreadsheetRaw/SheetRaw";
import {
  emptyTableStateIdentified,
  type FetchTargetIdentified,
  type TableStateIdentified,
} from "../ClassTypes/StateIdentified";
import {
  SpreadsheetBaseIdentified,
  type SpreadsheetIdentifiedProps,
} from "./SpreadsheetBaseIdentified";

export type TableIdentifiedProps = SpreadsheetIdentifiedProps & {
  tableId: string;
};

// A Table deleted and inserted again carries a new ID, so it reads as missing until the configs are regenerated.
export function managedTableAddress<TN extends TableName>(
  table: TableSchema<TN>,
): { tableId: string } {
  return { tableId: table.tableId };
}

export class TableBaseIdentified extends SpreadsheetBaseIdentified {
  readonly sheetGid: number;
  readonly knownTableId: string;
  constructor({
    spreadsheetStateRaw,
    spreadsheetStateIdentified,
    feedbackColumnIds,
    tableId,
  }: TableIdentifiedProps) {
    super({
      spreadsheetStateRaw,
      spreadsheetStateIdentified,
      feedbackColumnIds,
    });
    this.knownTableId = tableId;
    this.sheetGid = this.rawTable.sheetGid;
    if (!this.tablesStateIdentified.has(tableId)) {
      this.tablesStateIdentified.set(tableId, emptyTableStateIdentified());
    }
  }
  get tableIdentifiedProps(): TableIdentifiedProps {
    return {
      ...this.spreadsheetIdentifiedProps,
      tableId: this.knownTableId,
    };
  }
  private get rawTable(): TableBaseRaw {
    return new TableBaseRaw({
      ...this.spreadsheetRawProps,
      tableId: this.knownTableId,
    });
  }
  private get rawSheet(): SheetRaw {
    return new SheetRaw({
      ...this.spreadsheetRawProps,
      sheetGid: this.sheetGid,
    });
  }
  protected get tableState(): TableStateIdentified {
    return Val.assert(
      this.tablesStateIdentified.get(this.knownTableId),
      `Identified Table state for tableId ${this.knownTableId}`,
    );
  }
  get fetchTargets(): FetchTargetIdentified[] {
    return this.tableState.fetchQueue.targets;
  }
  // Rule and protection fetches still need the Table's column IDs.
  get isPreppedToFetch(): boolean {
    return this.fetchTargets.length > 0 || this.rawSheet.hasGatheredFetch;
  }
  // A Table its configs record is described by its own entry, since its sheet may record several.
  protected get tableSchema(): TableSchema {
    if (tableConfigsByTableId().has(this.knownTableId)) {
      return TableSchema.fromTableId(this.knownTableId);
    }
    return TableSchema.fromSheetGid(this.sheetGid);
  }
  clearFetchTargets(): void {
    this.tableState.fetchQueue.targets = [];
  }
}
