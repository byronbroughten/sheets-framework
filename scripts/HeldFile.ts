// The entered-data gate's terminal half, shared by sync and chores: the --confirm token and the dry run's held list on disk.
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join, relative } from "node:path";

import type { SyncHeld } from "../src/01_SpreadsheetSchema/configReaders/syncRunResult.ts";
import type {
  ConfirmedList,
  HeldList,
} from "../src/nodeHost/EnteredDataGate.ts";
import type { HeldItem } from "../src/nodeHost/FlushClassifier/flushClassifierTypes.ts";

// Each command keeps its own held file, so a chore's dry run never covers a sync's send.
const heldFileNames = {
  sync: "sync-held.json",
  chore: "chore-held.json",
} as const;

export type HeldFileCommand = keyof typeof heldFileNames;

interface HeldFileContents {
  target: string;
  token: string;
  flushes: HeldItem[][];
}

export interface HeldFileProps {
  packageDir: string;
  command: HeldFileCommand;
  isSending: boolean;
  confirmToken: string | undefined;
}

// The dry run's held list on disk, which a send with --confirm reads back.
export class HeldFile {
  readonly path: string;
  readonly isSending: boolean;
  readonly confirmToken: string | undefined;
  constructor({ packageDir, command, isSending, confirmToken }: HeldFileProps) {
    // In the package's .probe/, beside the probe's output, since it holds live cell values and is never committed.
    this.path = join(packageDir, ".probe", heldFileNames[command]);
    this.isSending = isSending;
    this.confirmToken = confirmToken;
  }
  static init(props: HeldFileProps): HeldFile {
    return new HeldFile(props);
  }
  // Undefined when the file is missing or isn't a held list; the gate compares each entry with the flush it judges.
  get flushes(): unknown[][] | undefined {
    if (!existsSync(this.path)) return undefined;
    try {
      const { flushes }: { flushes?: unknown } = JSON.parse(
        readFileSync(this.path, "utf8"),
      );
      if (!Array.isArray(flushes)) return undefined;
      if (!flushes.every((flush): flush is unknown[] => Array.isArray(flush))) {
        return undefined;
      }
      return flushes;
    } catch {
      return undefined;
    }
  }
  // A send's token, with the dry run's held list it covers; a missing or unreadable held file leaves the list undefined, so the gate counts it mismatched.
  get confirmedList(): ConfirmedList | undefined {
    const { isSending, confirmToken } = this;
    if (!isSending || confirmToken === undefined) return undefined;
    return { token: confirmToken, flushes: this.flushes };
  }
  // Writes a dry run's held items with their values to the held file; undefined, writing nothing, when no flush held any.
  async dryRunHeld(
    { token, flushes }: HeldList,
    target: string,
  ): Promise<SyncHeld | undefined> {
    if (flushes.every((items) => items.length === 0)) return undefined;
    const { heldGroups } = await import("../src/nodeHost/heldGroups.ts");
    this._write({ target, token, flushes });
    return {
      groups: heldGroups(flushes),
      token,
      fullListPath: relative(process.cwd(), this.path),
    };
  }
  _write(contents: HeldFileContents): void {
    mkdirSync(dirname(this.path), { recursive: true });
    writeFileSync(this.path, `${JSON.stringify(contents, null, 2)}\n`);
  }
}

// The token after `--confirm`, which the dry run printed beside its held list.
export function parseConfirmToken(argv: string[]): string | undefined {
  const index = argv.indexOf("--confirm");
  if (index === -1) return undefined;
  const token = argv[index + 1];
  if (token === undefined || token.startsWith("--")) {
    throw new Error("--confirm takes the token the dry run printed.");
  }
  return token;
}
