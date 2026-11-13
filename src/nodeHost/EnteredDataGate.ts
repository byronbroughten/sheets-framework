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

export interface EnteredDataGateBinding {
  target: string;
  // The dry run's held list, loaded back with the token the send was given; undefined when it was given none.
  confirm: HeldList | undefined;
  // Injected by the Node half, since src/ takes no Node API.
  hash: (input: string) => string;
}

export interface EnteredDataGateProps extends EnteredDataGateBinding {
  isDryRun: boolean;
}

// ADR 0034: holds each flush the classifier holds until a hash-bound token covers it; a dry run withholds every flush.
export class EnteredDataGate {
  private target: string;
  private confirm: HeldList | undefined;
  private hash: (input: string) => string;
  private isDryRun: boolean;
  private reads: SheetsHttpRead[];
  private flushes: HeldItem[][];
  private landedFlushes: number;
  constructor(props: EnteredDataGateProps) {
    this.target = props.target;
    this.confirm = props.confirm;
    this.hash = props.hash;
    this.isDryRun = props.isDryRun;
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
  recordRead(read: SheetsHttpRead): void {
    this.reads.push(read);
  }
  // The reads restart at every flush, withheld or sent, so a dry run judges each flush as its send will.
  passFlush(
    requests: GoogleRequest[],
    send: () => GoogleBatchUpdateResponse,
  ): GoogleBatchUpdateResponse {
    const flushIndex = this.flushes.length;
    const items = FlushClassifier.init({
      reads: this.reads,
      requests,
      dropdownStyleReset: "ask",
      forcedTraits: [],
    }).heldItems();
    this.flushes.push(items);
    this.reads = [];
    if (this.isDryRun) return {};
    this._assertCovered(flushIndex, items);
    const response = send();
    this.landedFlushes += 1;
    return response;
  }
  private _assertCovered(flushIndex: number, items: HeldItem[]): void {
    if (items.length === 0) return;
    if (this.confirm === undefined) {
      throw new HeldFlushRefusal("uncovered", this.landedFlushes);
    }
    const { token, flushes } = this.confirm;
    const isCovered =
      token === this._tokenFor(flushes) &&
      JSON.stringify(flushes[flushIndex]) === JSON.stringify(items);
    if (!isCovered) {
      throw new HeldFlushRefusal("mismatched", this.landedFlushes);
    }
  }
  private _tokenFor(flushes: HeldItem[][]): string {
    return this.hash(JSON.stringify({ target: this.target, flushes }));
  }
}
