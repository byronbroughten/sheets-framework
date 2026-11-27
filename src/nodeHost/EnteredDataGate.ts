import type {
  GoogleBatchUpdateResponse,
  GoogleRequest,
  SheetsHttpRead,
} from "../00_Source/GoogleSheets/GoogleSheetsAPI.ts";
import { HeldFlushRefusal } from "../00_Source/RawSource/HeldFlushRefusal.ts";
import {
  defaultDropdownStyleReset,
  type DropdownStyleReset,
} from "../00_Source/RawSource/RawSource.ts";
import type { ForcedTrait } from "../01_SpreadsheetSchema/configReaders/syncRunResult.ts";
import { FlushClassifier } from "./FlushClassifier.ts";
import type { HeldItem } from "./FlushClassifier/flushClassifierTypes.ts";
import {
  type JudgedReads,
  judgedReadsOf,
} from "./FlushClassifier/judgedReads.ts";
import {
  type FlushSnapshotPlan,
  flushSnapshotPlan,
} from "./flushSnapshotPlan.ts";
import type { HeldFlushes } from "./heldGroups.ts";

// Every flush's held items in flush order, a free flush's empty list included, and the token binding them to a target, which hashes only the flushes.
export interface HeldList extends HeldFlushes {
  token: string;
}

// The token a send was given, and the dry run's flushes loaded back from its held file; undefined flushes when that file was missing or unreadable.
// Unknown items, since the gate only hashes and compares them, so an edited file reads as mismatched.
export interface ConfirmedList {
  token: string;
  flushes: unknown[][] | undefined;
}

export interface EnteredDataGateProps {
  target: string;
  // Undefined when the send was given no token.
  confirmedList: ConfirmedList | undefined;
  // Undefined when the app sets none.
  dropdownStyleReset: DropdownStyleReset | undefined;
  // Injected by the Node half, since src/ takes no Node API.
  hash: (input: string) => string;
  // What a sync --force forces; only the send snapshot reads them, so a forced send needs no token.
  forcedTraits?: ForcedTrait[];
}

// ADR 0034: holds each flush the classifier holds until a hash-bound token covers it; the host withholds a dry run's flushes.
export class EnteredDataGate {
  private target: string;
  private confirmedList: ConfirmedList | undefined;
  private dropdownStyleReset: DropdownStyleReset;
  private hash: (input: string) => string;
  private forcedTraits: ForcedTrait[];
  private reads: SheetsHttpRead[];
  private judgedReads: JudgedReads;
  private flushes: HeldItem[][];
  private versionHistoryItems: HeldItem[][];
  private landedFlushes: number;
  private latestSnapshotPlan: FlushSnapshotPlan;
  constructor(props: EnteredDataGateProps) {
    this.target = props.target;
    this.confirmedList = props.confirmedList;
    this.dropdownStyleReset =
      props.dropdownStyleReset ?? defaultDropdownStyleReset;
    this.hash = props.hash;
    this.forcedTraits = props.forcedTraits ?? [];
    this.reads = [];
    this.judgedReads = judgedReadsOf([]);
    this.flushes = [];
    this.versionHistoryItems = [];
    this.landedFlushes = 0;
    this.latestSnapshotPlan = { columns: [], versionHistoryItems: [] };
  }
  static init(props: EnteredDataGateProps): EnteredDataGate {
    return new EnteredDataGate(props);
  }
  get heldList(): HeldList {
    return {
      token: this._tokenFor(this.flushes),
      flushes: this.flushes,
      versionHistoryItems: this.versionHistoryItems,
    };
  }
  // What the send snapshot reads before the flush classify last judged goes out.
  get snapshotPlan(): FlushSnapshotPlan {
    return this.latestSnapshotPlan;
  }
  integrateRead(read: SheetsHttpRead): void {
    this.reads.push(read);
  }
  // A flush with no reads of its own is judged against the last ones, as docs/how-it-runs/sync.md's entered-data gate says.
  classify(requests: GoogleRequest[]): this {
    if (this.reads.length > 0) {
      this.judgedReads = judgedReadsOf(this.reads);
      this.reads = [];
    }
    const { judgedReads } = this;
    const hasShiftedCells = judgedReads.tracker.hasShiftedCells;
    const items = FlushClassifier.init({
      judgedReads,
      requests,
      dropdownStyleReset: this.dropdownStyleReset,
      forcedTraits: [],
    }).heldItems();
    this.latestSnapshotPlan = flushSnapshotPlan({
      reads: judgedReads.reads,
      items,
      forcedTraits: this.flushes.length === 0 ? this.forcedTraits : [],
      hasShiftedCells,
    });
    this.flushes.push(items);
    this.versionHistoryItems.push(this.latestSnapshotPlan.versionHistoryItems);
    return this;
  }
  // Sends the flush classify last judged, once its held items are covered.
  flush(send: () => GoogleBatchUpdateResponse): GoogleBatchUpdateResponse {
    this._validateCovered();
    const response = send();
    this.landedFlushes += 1;
    return response;
  }
  private _validateCovered(): void {
    const flushIndex = this.flushes.length - 1;
    const items = this.flushes[flushIndex] ?? [];
    if (items.length === 0) return;
    if (this.confirmedList === undefined) {
      throw new HeldFlushRefusal({
        kind: "uncovered",
        landedFlushes: this.landedFlushes,
        flushCount: undefined,
      });
    }
    const { token, flushes } = this.confirmedList;
    const isCovered =
      flushes !== undefined &&
      token === this._tokenFor(flushes) &&
      JSON.stringify(flushes[flushIndex]) === JSON.stringify(items);
    if (!isCovered) {
      throw new HeldFlushRefusal({
        kind: "mismatched",
        landedFlushes: this.landedFlushes,
        flushCount: flushes?.length,
      });
    }
  }
  private _tokenFor(flushes: unknown[][]): string {
    return this.hash(JSON.stringify({ target: this.target, flushes }));
  }
}
