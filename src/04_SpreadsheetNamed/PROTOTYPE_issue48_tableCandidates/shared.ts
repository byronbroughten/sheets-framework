// PROTOTYPE for sheets-framework#48: throwaway and type-only. It never reaches master.
// A trimmed copy of the app's schema, so each candidate file can rewrite real call sites
// from sheets-real-estate@18b84eb and `tsc` can check them. Nothing here runs.
// Members keep today's names, so the candidate files differ only in their nouns.
// Candidates: sheetPrimary.ts (i), tablePrimary.ts (ii), tableParts.ts (iii).
import type { RowIdByName } from "../Types/RowIdByName";

export type { RowIdByName };
export type SerialDate = number & { readonly __serialDate: true };
export type ActionReturn = string;

export interface Tables {
  occupancy: {
    id: string;
    latestOccupancyTermsId: string;
    nextTermsNoticeSentDate: SerialDate;
    nextTermsStartDate: SerialDate;
    nextTermsEndDate: SerialDate | "";
    nextBaseRentChargeMonthly: number;
    nextTermsNotes: string;
    updateTermsSelect: boolean;
  };
  occupancyTerms: {
    id: string;
    occupancyId: string;
    noticeDate: SerialDate;
    startDate: SerialDate;
    endDate: SerialDate | "";
    rentChargeMonthly: number;
    notes: string;
  };
  occCharge: { id: string; occupancyId: string; amount: number };
  occChargeReduce: { id: string; chargeId: string; amount: number };
  property: { id: string; name: string };
  unit: { id: string; name: string; propertyId: string };
  variable: { occupancyLedgerOccId: string; occupancyLedgerDateRan: SerialDate };
  occupancyLedger: {
    date: SerialDate;
    description: string;
    charge: number | "";
    payment: number | "";
  };
}

export type TableName = keyof Tables;
export type ColName<TN extends TableName> = keyof Tables[TN] & string;
export type Vals<TN extends TableName> = Tables[TN];
export type LedgerLine = Vals<"occupancyLedger">;

// The same row in all three candidates; rows aren't in question here.
export interface RowNamed<TN extends TableName> {
  readonly rowIndex: number;
  readonly isBlank: boolean;
  value<CN extends ColName<TN>>(col: CN): Vals<TN>[CN];
  valueOrEmpty<CN extends ColName<TN>>(col: CN): Vals<TN>[CN] | "";
  updateValue<CN extends ColName<TN>>(col: CN, value: Vals<TN>[CN]): void;
  updateValues(values: Partial<Vals<TN>>): void;
  delete(): void;
}

// What the #47 inventory found the app calling on today's SheetNamed; every candidate hangs these on a body.
export interface BodyMembers<TN extends TableName> {
  prepFetchColumnsFull(...cols: ColName<TN>[]): void;
  prepFetchColumnsSpecific(rowIndexes: number[], ...cols: ColName<TN>[]): void;
  prepFetchRowIdAndName(): void;
  row(rowIndex: number): RowNamed<TN>;
  rowByValue<CN extends ColName<TN>>(col: CN, value: Vals<TN>[CN]): RowNamed<TN>;
  rowsFiltered(filter: Partial<Vals<TN>>): RowNamed<TN>[];
  readonly rowIndexesActiveWithData: number[];
  rowIdByName(name: string): RowIdByName;
  readonly topRow: RowNamed<TN>;
  appendRowWithAllVals(values: Vals<TN>): void;
  DELETE_ALL_DATA_ROWS(): void;
}

// A Table column, held as a value: what a formula rewrite (#49) or a dropdown validator source needs.
export interface TableColumnNamed<TN extends TableName, CN extends ColName<TN>> {
  readonly tableName: TN;
  readonly colName: CN;
  readonly headerText: string;
  readonly tableReference: string;
}

export interface SpreadsheetCommon {
  fetchAllPrepped(): void;
  today(): SerialDate;
  dayBefore(date: SerialDate): SerialDate;
}
