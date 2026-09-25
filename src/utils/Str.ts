export type RemoveFirstN<
  T extends string,
  N extends number,
  ARR extends unknown[] = [],
> = ARR["length"] extends N
  ? T
  : T extends `${string}${infer Rest}`
    ? RemoveFirstN<Rest, N, [...ARR, unknown]>
    : T;

export type TakeFirstN<
  S extends string,
  N extends number,
  Acc extends unknown[] = [],
> = Acc["length"] extends N
  ? ""
  : S extends `${infer First}${infer Rest}`
    ? `${First}${TakeFirstN<Rest, N, [...Acc, unknown]>}`
    : S;

export type CombineStrings<S1 extends string, S2 extends string> = `${S1}${S2}`;
export type TextJoin<
  S1 extends string,
  S2 extends string,
  D extends string,
> = `${S1}${D}${S2}`;

type Digit = "0" | "1" | "2" | "3" | "4" | "5" | "6" | "7" | "8" | "9";
type LowerAlpha =
  | "a"
  | "b"
  | "c"
  | "d"
  | "e"
  | "f"
  | "g"
  | "h"
  | "i"
  | "j"
  | "k"
  | "l"
  | "m"
  | "n"
  | "o"
  | "p"
  | "q"
  | "r"
  | "s"
  | "t"
  | "u"
  | "v"
  | "w"
  | "x"
  | "y"
  | "z";
type AlphaNumericChar = Digit | LowerAlpha;

type RemoveChar<
  S extends string,
  C extends string,
> = S extends `${infer Before}${C}${infer After}`
  ? RemoveChar<`${Before}${After}`, C>
  : S;

type RemoveApostrophes<S extends string> = RemoveChar<RemoveChar<S, "'">, "’">;

type SplitWords<
  S extends string,
  Current extends string = "",
  Words extends string[] = [],
> = S extends `${infer C}${infer Rest}`
  ? C extends AlphaNumericChar
    ? SplitWords<Rest, `${Current}${C}`, Words>
    : SplitWords<Rest, "", Current extends "" ? Words : [...Words, Current]>
  : Current extends ""
    ? Words
    : [...Words, Current];

type CapitalizeWord<S extends string> = S extends `${infer F}${infer R}`
  ? `${Uppercase<F>}${R}`
  : S;

type JoinCamelWords<
  Words extends string[],
  IsFirst extends boolean = true,
> = Words extends [infer Head extends string, ...infer Rest extends string[]]
  ? IsFirst extends true
    ? `${Head}${JoinCamelWords<Rest, false>}`
    : `${CapitalizeWord<Head>}${JoinCamelWords<Rest, false>}`
  : "";

// Type-level mirror of Str.sentenceToCamelCase: splits on runs of
// non-alphanumeric characters (apostrophes removed rather than treated as a
// split point, matching the runtime regex behavior) and camelCases the result.
export type SentenceToCamelCase<S extends string> = JoinCamelWords<
  SplitWords<RemoveApostrophes<Lowercase<S>>>
>;

export const Str = {
  combineStrings: <S1 extends string, S2 extends string>(
    str1: S1,
    str2: S2,
  ): CombineStrings<S1, S2> => {
    return `${str1}${str2}` as CombineStrings<S1, S2>;
  },
  removeFirstN<T extends string, N extends number>(
    str: T,
    n: N,
  ): RemoveFirstN<T, N> {
    return str.split("").slice(n).join("") as RemoveFirstN<T, N>;
  },
  takeFirstN<T extends string, N extends number>(
    str: T,
    n: N,
  ): TakeFirstN<T, N> {
    return str.split("").slice(0, n).join("") as TakeFirstN<T, N>;
  },
  sentenceToCamelCase<S extends string>(sentence: S): SentenceToCamelCase<S> {
    /**
     * Converts a header sentence (e.g. "Column ID") into a camelCase key
     * (e.g. "columnId"), used to robustly match row-3 headers regardless of
     * minor spacing/punctuation/capitalization differences.
     */
    return sentence
      .toLowerCase()
      .trim()
      .replace(/['’]/g, "") // remove straight & curly apostrophes
      .split(/[^a-zA-Z0-9]+/)
      .filter(Boolean)
      .map((word, index) => {
        if (index === 0) return word;
        return word.charAt(0).toUpperCase() + word.slice(1);
      })
      .join("") as SentenceToCamelCase<S>;
  },
};
