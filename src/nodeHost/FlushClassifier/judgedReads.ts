import type { FlushRead } from "./flushClassifierTypes.ts";
import { FlushTracker } from "./FlushTracker.ts";
import { RunReads } from "./RunReads.ts";

// The reads flushes are judged against, and the grid the flushes judged against them since left.
export interface JudgedReads {
  reads: RunReads;
  tracker: FlushTracker;
}

export function judgedReadsOf(reads: FlushRead[]): JudgedReads {
  const runReads = new RunReads(reads);
  return { reads: runReads, tracker: new FlushTracker(runReads.gridSizes) };
}
