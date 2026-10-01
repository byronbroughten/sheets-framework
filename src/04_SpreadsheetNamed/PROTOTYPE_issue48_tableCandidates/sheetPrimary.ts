// PROTOTYPE #48, candidate (i): Sheet stays primary and the Table is its child, `sheet.table`.
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
  sheet<SN extends TableName>(sheetName: SN): SheetNamed<SN>;
}

// The generic is still a sheet name, but every column union under it belongs to the sheet's one Table.
interface SheetNamed<SN extends TableName> {
  readonly title: string;
  readonly table: TableNamed<SN>;
}

interface TableNamed<SN extends TableName> extends BodyMembers<SN> {
  readonly name: string;
  readonly sheet: SheetNamed<SN>;
  column<CN extends ColName<SN>>(col: CN): TableColumnNamed<SN, CN>;
}

// updateTerms.ts:7-65; selectedRowIndexes stay sheet-row indexes, as the dispatch hands them over.
export function updateTerms(
  ss: SpreadsheetNamed,
  selectedRowIndexes: number[],
): ActionReturn {
  ss.sheet("occupancy").table.prepFetchColumnsSpecific(
    selectedRowIndexes,
    "id",
    "latestOccupancyTermsId",
    "nextTermsNoticeSentDate",
    "nextTermsStartDate",
    "nextTermsEndDate",
    "nextBaseRentChargeMonthly",
    "nextTermsNotes",
  );
  ss.sheet("occupancyTerms").table.prepFetchColumnsFull(
    "id",
    "startDate",
    "endDate",
  );
  ss.fetchAllPrepped();
  const occupancy = ss.sheet("occupancy").table;
  const occupancyTerms = ss.sheet("occupancyTerms").table;
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
  const unit = ss.sheet("unit").table.rowIdByName(unitName);
  if (unit.found !== "one") return undefined;
  return ss.sheet("unit").table.row(unit.rowIndex).value("propertyId");
}

// PEO:223: the one genuinely sheet-scoped read; here it stays one hop away.
export function unresolved(
  ss: SpreadsheetNamed,
  match: Exclude<RowIdByName, { found: "one" }>,
  sheetName: "unit" | "property",
  name: string,
): string {
  const { title } = ss.sheet(sheetName);
  if (match.found === "many") {
    return `${match.rowCount} rows of ${title} are named "${name}"`;
  }
  return `no row of ${title} is named "${name}"`;
}

// PEO:265: the header text, today from config through `row.cell(col).schema`.
export function blankComplaint(ss: SpreadsheetNamed): string {
  return `fill in ${ss.sheet("unit").table.column("propertyId").headerText}`;
}

// OLO:131, 155-161 and 177-180.
export function rebuildLedger(
  ss: SpreadsheetNamed,
  occupancyId: string,
  lines: LedgerLine[],
): Map<string, RowNamed<"occCharge">> {
  ss.sheet("variable").table.topRow.updateValues({
    occupancyLedgerOccId: occupancyId,
    occupancyLedgerDateRan: ss.today(),
  });
  const charges = ss.sheet("occCharge").table;
  const chargesById = charges.rowIndexesActiveWithData.reduce(
    (byId, rowIndex) => {
      const charge = charges.row(rowIndex);
      byId.set(charge.valueOrEmpty("id"), charge);
      return byId;
    },
    new Map<string, RowNamed<"occCharge">>(),
  );
  const ledger = ss.sheet("occupancyLedger").table;
  ledger.DELETE_ALL_DATA_ROWS();
  lines.forEach((line) => ledger.appendRowWithAllVals(line));
  return chargesById;
}

// #49's need and a future validator source: both start from a Table, but here are reached through its sheet.
export function occupancyIdReference(ss: SpreadsheetNamed): string {
  return ss.sheet("occCharge").table.column("occupancyId").tableReference;
}
