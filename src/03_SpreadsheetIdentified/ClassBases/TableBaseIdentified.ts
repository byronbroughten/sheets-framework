import { Val } from "@byronbroughten/utils/val";

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

// Always the recorded ID: a Table deleted and inserted again reads as missing until the configs are regenerated.
// `sheetGid` rides along so the Raw Tables built from these props are placed before their properties arrive.
export interface TableAddressIdentified {
  tableId: string;
  sheetGid?: number;
}

export type TableIdentifiedProps = SpreadsheetIdentifiedProps &
  TableAddressIdentified;

export class TableBaseIdentified extends SpreadsheetBaseIdentified {
  readonly sheetGid: number;
  readonly tableId: string;
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
    this.tableId = tableId;
    this.sheetGid = TableSchema.fromTableId(tableId).sheetGid;
    if (!this.tablesStateIdentified.has(tableId)) {
      this.tablesStateIdentified.set(tableId, emptyTableStateIdentified());
    }
  }
  get tableIdentifiedProps(): TableIdentifiedProps {
    return {
      ...this.spreadsheetIdentifiedProps,
      tableId: this.tableId,
      sheetGid: this.sheetGid,
    };
  }
  private get rawTable(): TableBaseRaw {
    return new TableBaseRaw({
      ...this.spreadsheetRawProps,
      tableId: this.tableId,
      sheetGid: this.sheetGid,
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
      this.tablesStateIdentified.get(this.tableId),
      `Identified Table state for tableId ${this.tableId}`,
    );
  }
  get fetchTargets(): FetchTargetIdentified[] {
    return this.tableState.fetchQueue.targets;
  }
  // Rule and protection fetches still need the Table's column IDs.
  get isPreppedToFetch(): boolean {
    return this.fetchTargets.length > 0 || this.rawSheet.hasGatheredFetch;
  }
  // Described by its own entry, since its sheet may record several.
  protected get tableSchema(): TableSchema {
    return TableSchema.fromTableId(this.tableId);
  }
  clearFetchTargets(): void {
    this.tableState.fetchQueue.targets = [];
  }
}
