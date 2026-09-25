// tsc passing doesn't prove precision: a type that widens, or collapses to never, still compiles.
export type IsExactly<A, B> =
  (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2
    ? true
    : false;

export function assertType<T extends true>(assertion: T): T {
  return assertion;
}

export function assertNotType<T extends false>(assertion: T): T {
  return assertion;
}
