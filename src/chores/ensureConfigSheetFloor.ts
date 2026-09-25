import { ConfigCoordinator } from "../05_Operators/ConfigCoordinator";
import type { Chore } from "./Chore";

export const ensureConfigSheetFloor: Chore = {
  description:
    "Puts one whole-sheet edit warning on Spreadsheet Config, Sheet Config and Column Config, replaces a drifted floor warning, restores floor tab titles, Table names, headers, column IDs and group headings, and sets floor columns back to their seeded types.",
  action: (ss) => {
    const report = new ConfigCoordinator(
      ss.spreadsheetNamedProps,
    ).ensureConfigSheetFloor();
    ss.batchUpdateGSheets();
    return report;
  },
};
