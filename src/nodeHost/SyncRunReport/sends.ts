import type {
  SyncCreatedTable,
  SyncCreatedTableColumn,
  SyncLandedItem,
  SyncSendLine,
} from "../../01_SpreadsheetSchema/configReaders/syncRunResult";
import { reportText } from "../reportText";

export const sendLines = {
  of({ subject, what }: SyncSendLine): string {
    return `${reportText.subjectLabel(subject)}: ${what}`;
  },
  landed(item: SyncLandedItem): string {
    if (item.kind === "send") return sendLines.of(item);
    if (item.kind === "createdTable") return createdTableLine(item);
    const { table, header, column, columnId } = item;
    return `${table}: ${createdColumnText({ header, column, columnId })}.`;
  },
};

function createdColumnText({
  header,
  column,
  columnId,
}: SyncCreatedTableColumn): string {
  return `created "${header}" as ${column}, ${columnId}`;
}

function createdTableLine({
  table,
  tab,
  tableId,
  columns,
}: SyncCreatedTable): string {
  const columnTexts = columns.map(createdColumnText).join("; ");
  return `${tab}: created Table ${table}, ${tableId}; ${columnTexts}.`;
}
