import type { Merge } from "./merge";

export type UnionObj<Union extends string, P extends string, R> = {
  [K in Union]: Merge<Record<P, K>, R>;
}[Union];
