type OptionalPropertyNames<T> = {
  // eslint-disable-next-line @typescript-eslint/no-empty-object-type -- `{}` is the optional-property test
  [K in keyof T]-?: {} extends { [P in K]: T[K] } ? K : never;
}[keyof T];

type SpreadProperties<L, R, K extends keyof L & keyof R> = {
  [P in K]: L[P] | Exclude<R[P], undefined>;
};

type IdS<T> = T extends infer U ? { [K in keyof U]: U[K] } : never;

export type Merge<L, R> = IdS<
  Pick<L, Exclude<keyof L, keyof R>> &
    Pick<R, Exclude<keyof R, OptionalPropertyNames<R>>> &
    Pick<R, Exclude<OptionalPropertyNames<R>, keyof L>> &
    SpreadProperties<L, R, OptionalPropertyNames<R> & keyof L>
>;

export function merge<A extends object, B extends object>(
  a: A,
  b: B,
): Merge<A, B> {
  return { ...a, ...b } as unknown as Merge<A, B>;
}
