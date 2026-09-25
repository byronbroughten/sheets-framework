import type { SheetName } from "../01_SpreadsheetSchema/sheetConfigsTypes";
import type { Chore } from "./Chore";

export const addMissingColumnIds: Chore = {
  description:
    "Writes a generated column ID into every configured sheet's header cells that lack one.",
  action: (ss) => {
    const sheetNames = ss.schema.sheetNames;
    ss.fetchAllSheetProperties();
    sheetNames.forEach((sheetName) => {
      ss.sheetMeta(sheetName).uniformRow("columnId").prepFetchFull();
    });
    ss.fetchAllPrepped({ skipFetchingProperties: true });
    const addedBySheet = sheetNames.map((sheetName): [SheetName, number] => [
      sheetName,
      ss.sheetMeta(sheetName).addMissingColumnIds(),
    ]);
    ss.batchUpdateGSheets();
    return addedIdsSummary(addedBySheet);
  },
};

function addedIdsSummary(addedBySheet: [SheetName, number][]): string {
  const touched = addedBySheet.filter(([, added]) => added > 0);
  const total = touched.reduce((count, [, added]) => count + added, 0);
  if (total === 0) {
    return "Every configured column already has an ID.";
  }
  const perSheet = touched
    .map(([sheetName, added]) => `${sheetName} (${added})`)
    .join(", ");
  return `Added ${total} column ID(s) across ${touched.length} sheet(s): ${perSheet}.`;
}
