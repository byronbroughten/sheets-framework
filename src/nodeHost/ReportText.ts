import type { SyncSubject } from "../01_SpreadsheetSchema/configReaders/syncRunResult";

// The plain-text pieces every section of sync's and a chore's report is built from.
export const ReportText = {
  section(title: string, lines: string[]): string[] {
    if (lines.length === 0) return [];
    return ["", `== ${title} ==`, ...lines];
  },
  indent(depth: number, lines: string[]): string[] {
    return lines.map((line) => `${" ".repeat(depth)}${line}`);
  },
  formatCount(count: number): string {
    return count.toLocaleString("en-US");
  },
  counted(count: number, one: string, many: string): string {
    return `${ReportText.formatCount(count)} ${count === 1 ? one : many}`;
  },
  subjectLabel(subject: SyncSubject): string {
    if ("tab" in subject) {
      const gid = subject.gid === undefined ? "" : ` (gid ${subject.gid})`;
      const trait = subject.path === undefined ? "" : ` › ${subject.path}`;
      return `tab ${subject.tab}${gid}${trait}`;
    }
    const { table, column, rewrittenFrom, path } = subject;
    let columnLabel = column;
    if (column !== undefined && rewrittenFrom !== undefined) {
      const { wrote, columnId } = rewrittenFrom;
      columnLabel = `${column} (you wrote ${wrote}; ${columnId})`;
    }
    return [table, columnLabel, path]
      .filter((part) => part !== undefined)
      .join(" › ");
  },
};
