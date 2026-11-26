// A confirmed send's snapshot file, its undo line and the prune of old snapshots, which chore calls and sync is to call too; restore reads the file back. The caller names the folder; none is the package's .probe/snapshots.
import {
  mkdirSync,
  readdirSync,
  readFileSync,
  rmSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { join, relative, resolve } from "node:path";

import type { HeldItem } from "../src/nodeHost/FlushClassifier/flushClassifierTypes.ts";
import type { SendSnapshot } from "../src/nodeHost/SendSnapshotTaker.ts";

// Each takes a value, which a runner reading its command's name must skip.
export const snapshotFlags = {
  folder: "--snapshots",
  restoreScript: "--restore-script",
} as const;

const filePrefix = "snapshot-";
const keptDays = 30;
const dayMilliseconds = 24 * 60 * 60 * 1000;
// The package's own script, for a send whose caller names no restore.
const ownRestoreScript = "restore";

export interface SnapshotFolderProps {
  dir: string;
  restoreScript: string;
  now: () => Date;
}

export interface SnapshotFolderArgs {
  packageDir: string;
  argv: string[];
  now: () => Date;
}

// What one send leaves in the folder: the send's plan as it printed it, and what the send changed.
export interface SnapshotFile {
  target: string;
  command: string;
  plan: string;
  snapshot: SendSnapshot;
}

// The file as written, which restore reads back.
export interface SavedSnapshotFile extends SnapshotFile {
  takenAt: string;
}

type SnapshotFileName = Pick<
  SavedSnapshotFile,
  "target" | "command" | "takenAt"
>;

export class SnapshotFolder {
  readonly dir: string;
  readonly restoreScript: string;
  readonly now: () => Date;
  constructor({ dir, restoreScript, now }: SnapshotFolderProps) {
    this.dir = dir;
    this.restoreScript = restoreScript;
    this.now = now;
  }
  static init(props: SnapshotFolderProps): SnapshotFolder {
    return new SnapshotFolder(props);
  }
  // The folder and restore script a caller names in argv; a send naming none keeps its snapshot beside the held file.
  static fromArgv({
    packageDir,
    argv,
    now,
  }: SnapshotFolderArgs): SnapshotFolder {
    const named = flagValue(argv, snapshotFlags.folder);
    return SnapshotFolder.init({
      // Typed from where npm was invoked, as outputShownOf reads a path.
      dir:
        named === undefined
          ? join(packageDir, ".probe", "snapshots")
          : resolve(invokedFrom(), named),
      restoreScript:
        flagValue(argv, snapshotFlags.restoreScript) ?? ownRestoreScript,
      now,
    });
  }
  // Removes the snapshots older than 30 days by file time, so a seeded time ages a file with no wait.
  prune(): void {
    let names: string[];
    try {
      names = readdirSync(this.dir);
    } catch {
      return;
    }
    const cutoff = this.now().getTime() - keptDays * dayMilliseconds;
    for (const name of names) {
      if (!name.startsWith(filePrefix)) continue;
      const path = join(this.dir, name);
      if (statSync(path).mtimeMs < cutoff) rmSync(path);
    }
  }
  // Writes one file and returns its path; a file already there is never overwritten.
  write({ target, command, plan, snapshot }: SnapshotFile): string {
    const takenAt = this.now().toISOString();
    const path = join(this.dir, fileNameOf({ takenAt, target, command }));
    mkdirSync(this.dir, { recursive: true });
    const contents: SavedSnapshotFile = {
      target,
      command,
      takenAt,
      plan,
      snapshot,
    };
    writeFileSync(path, `${JSON.stringify(contents, null, 2)}\n`, {
      flag: "wx",
    });
    return path;
  }
  // The undo line, with the held items only version history undoes beside it.
  async undoLines(
    path: string,
    versionHistoryItems: HeldItem[],
  ): Promise<string[]> {
    const { GateReport } = await import("../src/nodeHost/GateReport.ts");
    return [
      `undo: npm run ${this.restoreScript} -- ${relative(invokedFrom(), path)}`,
      ...GateReport.versionHistoryLines(versionHistoryItems),
    ];
  }
}

// The snapshot file an undo line names, a path typed from where npm was invoked.
export function readSnapshotFile(named: string): SavedSnapshotFile {
  const path = resolve(invokedFrom(), named);
  let contents: unknown;
  try {
    contents = JSON.parse(readFileSync(path, "utf8"));
  } catch (error) {
    throw new Error(
      `Could not read the snapshot file ${named}: ${error.message}`,
    );
  }
  if (!isSavedSnapshotFile(contents)) {
    throw new Error(`${named} is not a send's snapshot file.`);
  }
  return contents;
}

// The fields restore reads; the file is one this module wrote.
function isSavedSnapshotFile(contents: unknown): contents is SavedSnapshotFile {
  const file = contents as Partial<SavedSnapshotFile> | null;
  return (
    typeof file?.target === "string" &&
    typeof file.command === "string" &&
    Array.isArray(file.snapshot?.columns) &&
    Array.isArray(file.snapshot.fingerprints)
  );
}

function flagValue(argv: string[], flag: string): string | undefined {
  const index = argv.indexOf(flag);
  if (index === -1) return undefined;
  const value = argv[index + 1];
  if (value === undefined || value.startsWith("--")) {
    throw new Error(`${flag} takes a value.`);
  }
  return value;
}

function invokedFrom(): string {
  return process.env.INIT_CWD ?? process.cwd();
}

function fileNameOf({ takenAt, target, command }: SnapshotFileName): string {
  const stamp = takenAt.replaceAll(":", "-");
  return `${filePrefix}${stamp}-${target}-${command.replaceAll(" ", "-")}.json`;
}
