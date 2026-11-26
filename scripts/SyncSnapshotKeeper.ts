import type { SyncRunResult } from "../src/01_SpreadsheetSchema/configReaders/syncRunResult.ts";
import type { NodeHost } from "../src/nodeHost/NodeHost.ts";
import type { SendSnapshot } from "../src/nodeHost/SendSnapshotTaker.ts";
import {
  printSnapshotNotice,
  type SnapshotFolder,
  type SnapshotFolderKeepArgs,
} from "./SnapshotFolder.ts";
import { labelled } from "./syncLabel.ts";

export interface SyncSnapshotKeeperProps {
  host: NodeHost;
  snapshotFolder: SnapshotFolder;
  target: string;
  isHeldSend: boolean;
}

// What a held send prints about its snapshot: the undo lines, or why it was lost.
export interface SnapshotNotice {
  text: string;
  isLost: boolean;
}

interface ReturnedSyncRun {
  result: SyncRunResult;
  plan: string;
}

export class SyncSnapshotKeeper {
  readonly host: NodeHost;
  readonly snapshotFolder: SnapshotFolder;
  readonly target: string;
  readonly isHeldSend: boolean;
  constructor({
    host,
    snapshotFolder,
    target,
    isHeldSend,
  }: SyncSnapshotKeeperProps) {
    this.host = host;
    this.snapshotFolder = snapshotFolder;
    this.target = target;
    this.isHeldSend = isHeldSend;
  }
  static init(props: SyncSnapshotKeeperProps): SyncSnapshotKeeper {
    return new SyncSnapshotKeeper(props);
  }
  async keepReturned({
    result,
    plan,
  }: ReturnedSyncRun): Promise<SnapshotNotice | undefined> {
    return this._keep({
      plan,
      takeSnapshot: () => this._landedSnapshot(result),
    });
  }
  // Flushes that landed before a throw still get their snapshot; with no report to print, the requests are the plan.
  async keepThrown(): Promise<void> {
    const plan = this.host.summary.lines.join("\n");
    const notice = await this._keep({
      plan,
      takeSnapshot: () => this.host.sendSnapshot(),
    });
    if (notice !== undefined) printSnapshotNotice(`${plan}\n\n${notice.text}`);
  }
  // A send stopped before any flush landed changed nothing, though a flush Google refused already had its snapshot read.
  private _landedSnapshot({
    outcome,
  }: SyncRunResult): SendSnapshot | undefined {
    if (outcome.kind === "stoppedBeforeSending") return undefined;
    return this.host.sendSnapshot();
  }
  // A confirmed or forced send prunes the old snapshots and writes its own; a forced send needs no token.
  private async _keep({
    plan,
    takeSnapshot,
  }: Pick<SnapshotFolderKeepArgs, "plan" | "takeSnapshot">): Promise<
    SnapshotNotice | undefined
  > {
    if (!this.isHeldSend) return undefined;
    try {
      const undo = await this.snapshotFolder.keep({
        takeSnapshot,
        target: this.target,
        command: "sync",
        plan,
      });
      if (undo === undefined) return undefined;
      return { text: undo, isLost: false };
    } catch (error) {
      // The send is already out, so a lost snapshot is reported and never takes the run's place.
      const text = labelled(
        `no snapshot was written, so only the spreadsheet's version history undoes this send: ${error.message}`,
      );
      return { text, isLost: true };
    }
  }
}
