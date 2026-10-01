// PROTOTYPE #48, candidate (iii): the Table splits into header row and body, as Excel JS slices it (#45).
// The Sheet is the same container as in (ii); only the Table's shape differs.
import type {
  ActionReturn,
  BodyMembers,
  ColName,
  LedgerLine,
  RowIdByName,
  RowNamed,
  SpreadsheetCommon,
  TableColumnNamed,
  TableName,
} from "./shared";

interface SpreadsheetNamed extends SpreadsheetCommon {
  table<TN extends TableName>(tableName: TN): TableNamed<TN>;
}

interface SheetNamed {
  readonly title: string;
  readonly tableNames: TableName[];
}

interface TableNamed<TN extends TableName> {
  readonly name: TN;
  readonly sheet: SheetNamed;
  readonly headerRow: TableHeaderRowNamed<TN>;
  readonly body: TableBodyNamed<TN>;
  column<CN extends ColName<TN>>(col: CN): TableColumnNamed<TN, CN>;
}

interface TableHeaderRowNamed<TN extends TableName> {
  readonly sheetRowIndex: number;
  text(col: ColName<TN>): string;
}

type TableBodyNamed<TN extends TableName> = BodyMembers<TN>;

// updateTerms.ts:7-65; selectedRowIndexes stay sheet-row indexes, as the dispatch hands them over.
export function updateTerms(
  ss: SpreadsheetNamed,
  selectedRowIndexes: number[],
): ActionReturn {
  ss.table("occupancy").body.prepFetchColumnsSpecific(
    selectedRowIndexes,
    "id",
    "latestOccupancyTermsId",
    "nextTermsNoticeSentDate",
    "nextTermsStartDate",
    "nextTermsEndDate",
    "nextBaseRentChargeMonthly",
    "nextTermsNotes",
  );
  ss.table("occupancyTerms").body.prepFetchColumnsFull(
    "id",
    "startDate",
    "endDate",
  );
  ss.fetchAllPrepped();
  const occupancy = ss.table("occupancy").body;
  const occupancyTerms = ss.table("occupancyTerms").body;
  selectedRowIndexes.forEach((rowIndex) => {
    const occRow = occupancy.row(rowIndex);
    const nextStartDate = occRow.value("nextTermsStartDate");
    const lastActiveTerm = occupancyTerms.rowByValue(
      "id",
      occRow.value("latestOccupancyTermsId"),
    );
    if (!lastActiveTerm.value("endDate")) {
      lastActiveTerm.updateValue("endDate", ss.dayBefore(nextStartDate));
    }
    occupancyTerms.appendRowWithAllVals({
      id: "",
      occupancyId: occRow.value("id"),
      noticeDate: occRow.value("nextTermsNoticeSentDate"),
      startDate: nextStartDate,
      endDate: occRow.value("nextTermsEndDate"),
      rentChargeMonthly: occRow.value("nextBaseRentChargeMonthly"),
      notes: occRow.value("nextTermsNotes"),
    });
  });
  return "Occupancy terms updated";
}

// PEO:150-176: a name lookup, then the row it found.
export function propertyIdOfUnit(
  ss: SpreadsheetNamed,
  unitName: string,
): string | undefined {
  const unit = ss.table("unit").body.rowIdByName(unitName);
  if (unit.found !== "one") return undefined;
  return ss.table("unit").body.row(unit.rowIndex).value("propertyId");
}

// PEO:223.
export function unresolved(
  ss: SpreadsheetNamed,
  match: Exclude<RowIdByName, { found: "one" }>,
  tableName: "unit" | "property",
  name: string,
): string {
  const { title } = ss.table(tableName).sheet;
  if (match.found === "many") {
    return `${match.rowCount} rows of ${title} are named "${name}"`;
  }
  return `no row of ${title} is named "${name}"`;
}

// PEO:265: the one call site the header part serves, and it could equally be `column(col).headerText`.
export function blankComplaint(ss: SpreadsheetNamed): string {
  return `fill in ${ss.table("unit").headerRow.text("propertyId")}`;
}

// OLO:131, 155-161 and 177-180.
export function rebuildLedger(
  ss: SpreadsheetNamed,
  occupancyId: string,
  lines: LedgerLine[],
): Map<string, RowNamed<"occCharge">> {
  ss.table("variable").body.topRow.updateValues({
    occupancyLedgerOccId: occupancyId,
    occupancyLedgerDateRan: ss.today(),
  });
  const charges = ss.table("occCharge").body;
  const chargesById = charges.rowIndexesActiveWithData.reduce(
    (byId, rowIndex) => {
      const charge = charges.row(rowIndex);
      byId.set(charge.valueOrEmpty("id"), charge);
      return byId;
    },
    new Map<string, RowNamed<"occCharge">>(),
  );
  const ledger = ss.table("occupancyLedger").body;
  ledger.DELETE_ALL_DATA_ROWS();
  lines.forEach((line) => ledger.appendRowWithAllVals(line));
  return chargesById;
}

export function occupancyIdReference(ss: SpreadsheetNamed): string {
  return ss.table("occCharge").column("occupancyId").tableReference;
}
