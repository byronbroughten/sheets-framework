import {
  type TableAddressRaw,
  TableBaseRaw,
} from "../../02_SpreadsheetRaw/ClassBases/TableBaseRaw";
import { Val } from "../../utils/Val";
import {
  emptyTableStateIdentified,
  type FetchTargetIdentified,
  type TableStateIdentified,
} from "../ClassTypes/StateIdentified";
import {
  SpreadsheetBaseIdentified,
  type SpreadsheetIdentifiedProps,
} from "./SpreadsheetBaseIdentified";

export type TableIdentifiedProps = SpreadsheetIdentifiedProps & TableAddressRaw;

export class TableBaseIdentified extends SpreadsheetBaseIdentified {
  readonly sheetGid: number;
  private readonly tableAddress: TableAddressRaw;
  private readonly rawTable: TableBaseRaw;
  constructor({
    spreadsheetStateRaw,
    spreadsheetStateIdentified,
    feedbackColumnIds,
    ...tableAddress
  }: TableIdentifiedProps) {
    super({
      spreadsheetStateRaw,
      spreadsheetStateIdentified,
      feedbackColumnIds,
    });
    this.tableAddress = tableAddress;
    this.rawTable = new TableBaseRaw({
      ...this.spreadsheetRawProps,
      ...tableAddress,
    });
    this.sheetGid = this.rawTable.sheetGid;
    this._ensureTableState();
  }
  get tableIdentifiedProps(): TableIdentifiedProps {
    return {
      ...this.spreadsheetIdentifiedProps,
      ...this.tableAddress,
    };
  }
  private get tableStateBeforeProperties(): TableStateIdentified | undefined {
    return this.tableBeforePropertiesBySheet.get(this.sheetGid);
  }
  protected get tableState(): TableStateIdentified {
    return this._resolveTableState();
  }
  get fetchTargets(): FetchTargetIdentified[] {
    return this.tableState.fetchQueue.targets;
  }
  // Rule and protection fetches still need the Table's column IDs.
  get isPreppedToFetch(): boolean {
    return this.fetchTargets.length > 0 || this.rawTable.hasGatheredSheetFetch;
  }
  // Absent until the sheet's one Table is fetched.
  knownTableId(): string | undefined {
    if ("tableId" in this.tableAddress) return this.tableAddress.tableId;
    return this.rawTable.onlyTableId();
  }
  clearFetchTargets(): void {
    this.tableState.fetchQueue.targets = [];
  }
  // Before its Table is known, a handle reads the sheet's queue.
  private _resolveTableState(): TableStateIdentified {
    const tableId = this.knownTableId();
    const tableState =
      tableId === undefined
        ? undefined
        : this.tablesStateIdentified.get(tableId);
    return Val.assert(
      tableState ?? this.tableStateBeforeProperties,
      `Identified Table state for sheetGid ${this.sheetGid}`,
    );
  }
  private _ensureTableState(): void {
    const tableId = this.knownTableId();
    if (tableId === undefined) {
      this._ensureTableStateBeforeProperties();
      return;
    }
    if (!this.tablesStateIdentified.has(tableId)) {
      this.tablesStateIdentified.set(tableId, emptyTableStateIdentified());
    }
    if (this.rawTable.onlyTableId() === tableId) {
      this._adoptTableStateBeforeProperties(tableId);
    }
  }
  private _ensureTableStateBeforeProperties(): void {
    if (this.tableStateBeforeProperties !== undefined) return;
    this.tableBeforePropertiesBySheet.set(
      this.sheetGid,
      emptyTableStateIdentified(),
    );
  }
  // The sheet's one Table takes over what was prepped before it was known.
  private _adoptTableStateBeforeProperties(tableId: string): void {
    const adopted = this.tableStateBeforeProperties;
    if (adopted === undefined) return;
    const tableState = Val.assert(
      this.tablesStateIdentified.get(tableId),
      `Identified Table state for tableId ${tableId}`,
    );
    tableState.fetchQueue.targets.push(...adopted.fetchQueue.targets);
    this.tableBeforePropertiesBySheet.delete(this.sheetGid);
  }
}
