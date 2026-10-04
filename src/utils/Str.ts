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

type UpperAlpha = Uppercase<LowerAlpha>;
type Whitespace = " " | "\t" | "\n" | "\r";

type Trim<S extends string> = S extends `${Whitespace}${infer Rest}`
  ? Trim<Rest>
  : S extends `${infer Rest}${Whitespace}`
    ? Trim<Rest>
    : S;

type IsSentence<S extends string> =
  Trim<S> extends `${string}${Whitespace}${string}` ? true : false;

type StartsLowerAlpha<S extends string> = S extends `${LowerAlpha}${string}`
  ? true
  : false;

// Mirrors splitOnCase: a lone trailing s pluralizes the acronym.
type IsCaseBoundary<
  Prev extends string,
  C extends string,
  Rest extends string,
> = C extends UpperAlpha
  ? Prev extends LowerAlpha | Digit
    ? true
    : Prev extends UpperAlpha
      ? Rest extends `${infer Next}${infer After}`
        ? Next extends "s"
          ? StartsLowerAlpha<After>
          : StartsLowerAlpha<Next>
        : false
      : false
  : false;

type SplitTokenWords<
  S extends string,
  Prev extends string = "",
  Current extends string = "",
  Words extends string[] = [],
> = S extends `${infer C}${infer Rest}`
  ? C extends AlphaNumericChar | UpperAlpha
    ? IsCaseBoundary<Prev, C, Rest> extends true
      ? SplitTokenWords<Rest, C, Lowercase<C>, [...Words, Current]>
      : SplitTokenWords<Rest, C, `${Current}${Lowercase<C>}`, Words>
    : SplitTokenWords<
        Rest,
        C,
        "",
        Current extends "" ? Words : [...Words, Current]
      >
  : Current extends ""
    ? Words
    : [...Words, Current];

type SentenceOrTokenWords<S extends string> =
  IsSentence<S> extends true ? SplitWords<Lowercase<S>> : SplitTokenWords<S>;

// Mirrors Str.sentenceToCamelCase, so apostrophes are removed rather than split on.
export type SentenceToCamelCase<S extends string> = string extends S
  ? string
  : JoinCamelWords<SentenceOrTokenWords<RemoveApostrophes<S>>>;

export const Str = {
  combineStrings<S1 extends string, S2 extends string>(
    str1: S1,
    str2: S2,
  ): CombineStrings<S1, S2> {
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
  // A sentence splits only on non-alphanumerics, so "CapEx budget" keeps its key; one token also splits on case.
  words(text: string): string[] {
    const unpunctuated = text.trim().replace(/['’]/g, ""); // remove straight & curly apostrophes
    const isSentence = /\s/.test(unpunctuated);
    return (isSentence ? unpunctuated : splitOnCase(unpunctuated))
      .split(/[^a-zA-Z0-9]+/)
      .filter(Boolean)
      .map((word) => word.toLowerCase());
  },
  // Lets a header match despite spacing, punctuation or capitalization drift.
  sentenceToCamelCase<S extends string>(sentence: S): SentenceToCamelCase<S> {
    return Str.words(sentence)
      .map((word, index) => {
        if (index === 0) return word;
        return word.charAt(0).toUpperCase() + word.slice(1);
      })
      .join("") as SentenceToCamelCase<S>;
  },
};

function splitOnCase(token: string): string {
  return token
    .replace(/([a-z0-9])(?=[A-Z])/g, "$1 ")
    .replace(/([A-Z])(?=[A-Z][a-z])(?![A-Z]s(?![a-z]))/g, "$1 "); // a lone trailing s pluralizes the acronym
}
