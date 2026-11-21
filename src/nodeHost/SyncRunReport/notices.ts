import {
  fillRowIdsName,
  heldBackChangeWhat,
  type SyncNotice,
} from "../../01_SpreadsheetSchema/configReaders/syncRunResult.ts";
import { ReportText } from "../ReportText.ts";

export const noticeLines = {
  of(notice: SyncNotice): string {
    if (notice.kind === "text") return notice.text;
    if (notice.kind === "tabGone") {
      return `${notice.tab}: a person deleted this tab from the sheet, so sync dropped it from the configs.`;
    }
    if (notice.kind === "ruleHeldBack") {
      return `${ReportText.subjectLabel(notice.subject)}: ${heldBackChangeWhat(notice)}.`;
    }
    const { table } = notice;
    if (notice.kind === "tableGone") {
      return `${table}: a person deleted this Table from the sheet, so sync dropped it from the configs and retired its ID prefix.`;
    }
    return `${table}: ID is the first column. Sheets' Table name chip covers the cell above it, so ${table}'s ${fillRowIdsName} checkbox can't be clicked. Moving ID out of first place fixes it.`;
  },
};
