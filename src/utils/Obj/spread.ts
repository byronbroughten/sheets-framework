import type { Merge } from "./merge";

export type Spread<A extends readonly unknown[]> = A extends [
  infer L,
  ...infer R,
]
  ? Merge<L, Spread<R>>
  : unknown;

export function spread<A extends object[]>(...a: [...A]): Spread<A> {
  return Object.assign({}, ...a) as Spread<A>;
}
