import { Obj } from "@byronbroughten/utils/obj";
import { Val } from "@byronbroughten/utils/val";

import type { GoogleRequest } from "../00_Source/GoogleSheets/GoogleSheetsAPI";
import type { ForcedTrait } from "../05_Operators/syncRunResult";
import type {
  DropdownStyleReset,
  FlushRead,
  HeldItem,
} from "./FlushClassifier/flushClassifierTypes";
import { FlushTracker } from "./FlushClassifier/FlushTracker";
import { RequestJudge } from "./FlushClassifier/RequestJudge";
import {
  type ClassifiedVerb,
  requestJudgements,
} from "./FlushClassifier/requestJudgements";
import { RunReads } from "./FlushClassifier/RunReads";

export interface FlushClassifierProps {
  // Taken since the last flush landed: a stale read could free an overwrite.
  reads: FlushRead[];
  requests: GoogleRequest[];
  dropdownStyleReset: DropdownStyleReset;
  forcedTraits: ForcedTrait[];
}

interface RequestToJudge {
  request: GoogleRequest;
  requestIndex: number;
  reads: RunReads;
  tracker: FlushTracker;
}

// The entered-data gate's judge (ADR 0034): which of one flush's items are held, and why.
export class FlushClassifier {
  private flushReads: FlushRead[];
  private requests: GoogleRequest[];
  private dropdownStyleReset: DropdownStyleReset;
  private forcedTraits: ForcedTrait[];
  constructor(props: FlushClassifierProps) {
    this.flushReads = props.reads;
    this.requests = props.requests;
    this.dropdownStyleReset = props.dropdownStyleReset;
    this.forcedTraits = props.forcedTraits;
  }
  static init(props: FlushClassifierProps): FlushClassifier {
    return new FlushClassifier(props);
  }
  private get reads(): RunReads {
    return new RunReads(this.flushReads);
  }
  // Requests are judged in order, each against the grid the ones before it left.
  heldItems(): HeldItem[] {
    const reads = this.reads;
    const tracker = new FlushTracker(reads.gridSizes);
    const forced: HeldItem[] = this.forcedTraits.map((trait) => ({
      kind: "forcedTrait",
      trait,
    }));
    return this.requests.reduce((held, request, requestIndex) => {
      held.push(
        ...this._requestHeldItems({ request, requestIndex, reads, tracker }),
      );
      return held;
    }, forced);
  }
  private _requestHeldItems({
    request,
    requestIndex,
    reads,
    tracker,
  }: RequestToJudge): HeldItem[] {
    const verb = Obj.keys(request)[0];
    if (verb === undefined) return [];
    if (!Obj.isKey(requestJudgements, verb)) {
      return [{ kind: "unclassified", requestIndex, verb }];
    }
    const judge = new RequestJudge({
      reads,
      tracker,
      requestIndex,
      verb,
      dropdownStyleReset: this.dropdownStyleReset,
    });
    judgeRequest(verb, request, judge, tracker);
    return judge.heldItems;
  }
}

function judgeRequest<CV extends ClassifiedVerb>(
  verb: CV,
  request: GoogleRequest,
  judge: RequestJudge,
  tracker: FlushTracker,
): void {
  const inner = Val.assert<NonNullable<GoogleRequest[CV]>>(request[verb], verb);
  requestJudgements[verb](inner, judge, tracker);
}
