import type { Merge } from "./merge";

export type Spread<A extends readonly unknown[]> = A extends [
  infer L,
  ...infer R,
]
  ? Merge<L, Spread<R>>
  : unknown;

interface Sample1 {
  a: 1;
  b: 2;
}
interface Sample2 {
  a: 2;
  b: 2;
  c: 3;
}
type Test1 = Spread<[Sample1, Sample2]>;

interface Sample3 {
  a: 3;
  d: 4;
}
type Test3 = Spread<[Test1, Sample3]>;

function _spreadTest<T extends Test3>(_t: T): void {}
_spreadTest({
  b: 2,
  c: 3,
  a: 3,
  d: 4,
});

export function spread<A extends object[]>(...a: [...A]): Spread<A> {
  return Object.assign({}, ...a) as Spread<A>;
}
