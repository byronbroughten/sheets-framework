// `sheets-framework restore`: puts a send's snapshot back on the package's spreadsheet. See docs/how-it-runs/chore-dry-run.md.
import type { SheetsHttpTransport } from "../src/00_Source/GoogleSheets/GoogleSheetsAPI.ts";
import type { SnapshotRestore } from "../src/nodeHost/SendSnapshotRestorer.ts";
import type { ColumnLabel } from "../src/nodeHost/SendSnapshotRestorer/SnapshotColumnRestorer.ts";
import { liveTransport } from "./nodeHost.ts";
import type { SheetsConfig } from "./sheetsConfig.ts";
import {
  readSnapshotFile,
  type SavedSnapshotFile,
  snapshotTarget,
} from "./SnapshotFolder.ts";

const usage = `Usage: npm run restore -- <snapshot> [--force]
  <snapshot>  the snapshot file a send's undo line names
  --force     restore though rows were added, removed or reordered since the send`;

// A test's stand-in for Google.
export interface RestoreStandIn {
  transport: SheetsHttpTransport;
}

interface RestoreRunnerProps {
  sheetsConfig: SheetsConfig;
  snapshotPath: string | undefined;
  isForced: boolean;
  transport: SheetsHttpTransport;
}

export class RestoreRunner {
  readonly sheetsConfig: SheetsConfig;
  readonly snapshotPath: string | undefined;
  readonly isForced: boolean;
  readonly transport: SheetsHttpTransport;
  constructor({
    sheetsConfig,
    snapshotPath,
    isForced,
    transport,
  }: RestoreRunnerProps) {
    this.sheetsConfig = sheetsConfig;
    this.snapshotPath = snapshotPath;
    this.isForced = isForced;
    this.transport = transport;
  }
  static init(
    sheetsConfig: SheetsConfig,
    argv: string[],
    standIn?: RestoreStandIn,
  ): RestoreRunner {
    return new RestoreRunner({
      sheetsConfig,
      snapshotPath: argv.find((arg) => !arg.startsWith("--")),
      isForced: argv.includes("--force"),
      transport: standIn?.transport ?? liveTransport(),
    });
  }
  get target(): string {
    return snapshotTarget(this.sheetsConfig);
  }
  async run(): Promise<void> {
    if (this.snapshotPath === undefined) {
      console.log(usage);
      process.exitCode = 1;
      return;
    }
    const file = this._snapshotFile(this.snapshotPath);
    console.log(
      `restore: ${this.snapshotPath} — the send of ${file.command} at ${file.takenAt}\n`,
    );
    const { SendSnapshotRestorer } =
      await import("../src/nodeHost/SendSnapshotRestorer.ts");
    const restore = SendSnapshotRestorer.init({
      spreadsheetId: this.sheetsConfig.spreadsheetId,
      transport: this.transport,
      snapshot: file.snapshot,
      isForced: this.isForced,
    }).restore();
    console.log(restoreLines(restore).join("\n"));
    if (restore.kind === "stopped") process.exitCode = 1;
  }
  // A snapshot of another target's spreadsheet would land on the wrong one.
  _snapshotFile(path: string): SavedSnapshotFile {
    const file = readSnapshotFile(path);
    if (file.target !== this.target) {
      throw new Error(
        `${path} is a snapshot of ${file.target}'s spreadsheet, not ${this.target}'s.`,
      );
    }
    return file;
  }
}

function restoreLines(restore: SnapshotRestore): string[] {
  if (restore.kind === "stopped") {
    return [
      "RESTORE STOPPED — nothing was written. Rows were added, removed or reordered since the send on:",
      ...restore.changedTabs.map((tab) => `  ${tab}`),
      "Re-run with `-- --force` to restore anyway.",
    ];
  }
  const { changedTabs, restoredColumns, goneColumns, homelessRows } = restore;
  return [
    ...section(
      "restored anyway, though rows moved since the send, on:",
      changedTabs.map((tab) => `  ${tab}`),
    ),
    ...section(
      "restored:",
      restoredColumns.map(
        (column) => `  ${columnLabel(column)} (${column.rowCount} row(s))`,
      ),
    ),
    ...section(
      "not restored, its tab or Table is gone:",
      goneColumns.map(
        (column) => `  ${columnLabel(column)} (${column.columnId})`,
      ),
    ),
    ...section(
      "left to the spreadsheet's version history, with no row to land on:",
      homelessRows.map((row) => `  ${columnLabel(row)}, ${row.row}`),
    ),
  ];
}

function columnLabel({ tabTitle, header }: ColumnLabel): string {
  return `${tabTitle} › ${header}`;
}

function section(heading: string, lines: string[]): string[] {
  return lines.length === 0 ? [] : [heading, ...lines];
}

export async function runRestore(
  sheetsConfig: SheetsConfig,
  argv: string[],
): Promise<void> {
  await RestoreRunner.init(sheetsConfig, argv).run();
}
