import { Str } from "@byronbroughten/utils/str";

export interface IdPrefixLabel {
  label: string;
  idPrefix: string;
}

export interface RetiredIdPrefixesProps {
  previouslyRetired: readonly string[];
  previousByTableId: ReadonlyMap<string, string>;
  currentByTableId: ReadonlyMap<string, string>;
}

const vowels = new Set(["a", "e", "i", "o", "u"]);

export const idPrefixes = {
  // A retired prefix is taken for good: rows elsewhere may still carry IDs minted under it.
  fromTitle(
    title: string,
    prefixesInUse: ReadonlySet<string>,
    retiredPrefixes: readonly string[] = [],
  ): string {
    const taken = new Set([...prefixesInUse, ...retiredPrefixes]);
    const { base, remainingConsonants } = idPrefixBaseAndRemaining(title);
    if (!taken.has(base)) return base;
    let candidate = base;
    for (const letter of remainingConsonants) {
      candidate += letter;
      if (!taken.has(candidate)) return candidate;
    }
    for (let n = 2; ; n++) {
      const numbered = `${base}${n}`;
      if (!taken.has(numbered)) return numbered;
    }
  },
  retiredAfter({
    previouslyRetired,
    previousByTableId,
    currentByTableId,
  }: RetiredIdPrefixesProps): string[] {
    const current = new Set(currentByTableId.values());
    const retired = new Set(previouslyRetired);
    previousByTableId.forEach((idPrefix, tableId) => {
      if (!currentByTableId.has(tableId)) retired.add(idPrefix);
    });
    return [...retired].filter((idPrefix) => !current.has(idPrefix)).sort();
  },
  assertUnique(idPrefixLabels: ReadonlyArray<IdPrefixLabel>): void {
    const labelByPrefix = new Map<string, string>();
    idPrefixLabels.forEach(({ label, idPrefix }) => {
      if (!idPrefix) {
        throw new Error(`Sheet "${label}" has no ID prefix.`);
      }
      const existing = labelByPrefix.get(idPrefix);
      if (existing !== undefined) {
        throw new Error(
          `Sheets "${existing}" and "${label}" share ID prefix "${idPrefix}".`,
        );
      }
      labelByPrefix.set(idPrefix, label);
    });
  },
};

function idPrefixBaseAndRemaining(title: string): {
  base: string;
  remainingConsonants: string;
} {
  const titleWords = Str.words(title)
    .map((word) => word.replace(/[0-9]/g, ""))
    .filter((word) => word !== "");
  const lastWord = titleWords.at(-1);
  if (lastWord === undefined) return { base: "s", remainingConsonants: "" };
  const initials = titleWords.map((word) => word.charAt(0)).join("");
  const afterFirst = [...lastWord.slice(1)]
    .filter((letter) => !vowels.has(letter))
    .join("");
  if (initials.length >= 3) {
    return { base: initials, remainingConsonants: afterFirst };
  }
  const base = (initials + afterFirst).slice(0, 3);
  return {
    base,
    remainingConsonants: afterFirst.slice(base.length - initials.length),
  };
}
