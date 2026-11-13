import type {
  GoogleBatchUpdateResponse,
  GoogleRequest,
  SheetsHttpRead,
} from "../00_Source/GoogleSheets/GoogleSheetsAPI";
import { HeldFlushRefusal } from "../00_Source/RawSource/HeldFlushRefusal";
import { FlushClassifier } from "./FlushClassifier";
import type { HeldItem } from "./FlushClassifier/flushClassifierTypes";

// Every flush's held items in flush order, a free flush's empty list included, and the token that binds them to a target.
export interface HeldList {
  token: string;
  flushes: HeldItem[][];
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
  // Injected by the Node half, since src/ takes no Node API.
  hash: (input: string) => string;
}

// ADR 0034: holds each flush the classifier holds until a hash-bound token covers it; the host withholds a dry run's flushes.
export class EnteredDataGate {
  private target: string;
  private confirmedList: ConfirmedList | undefined;
  private hash: (input: string) => string;
  private reads: SheetsHttpRead[];
  private flushes: HeldItem[][];
  private landedFlushes: number;
  constructor(props: EnteredDataGateProps) {
    this.target = props.target;
    this.confirmedList = props.confirmedList;
    this.hash = props.hash;
    this.reads = [];
    this.flushes = [];
    this.landedFlushes = 0;
  }
  static init(props: EnteredDataGateProps): EnteredDataGate {
    return new EnteredDataGate(props);
  }
  get heldList(): HeldList {
    return {
      token: this._tokenFor(this.flushes),
      flushes: this.flushes,
    };
  }
  integrateRead(read: SheetsHttpRead): void {
    this.reads.push(read);
  }
  // The reads restart at every flush, withheld or sent, so a dry run judges each flush as its send will.
  classify(requests: GoogleRequest[]): this {
    this.flushes.push(
      FlushClassifier.init({
        reads: this.reads,
        requests,
        dropdownStyleReset: "ask",
        forcedTraits: [],
      }).heldItems(),
    );
    this.reads = [];
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
