export interface IdPrefixLabel {
  label: string;
  idPrefix: string;
}

const vowels = new Set(["a", "e", "i", "o", "u"]);

export const idPrefixes = {
  fromTitle(title: string, prefixesInUse: ReadonlySet<string>): string {
    const { base, remainingConsonants } = idPrefixBaseAndRemaining(title);
    if (!prefixesInUse.has(base)) return base;
    let candidate = base;
    for (const letter of remainingConsonants) {
      candidate += letter;
      if (!prefixesInUse.has(candidate)) return candidate;
    }
    for (let n = 2; ; n++) {
      const numbered = `${base}${n}`;
      if (!prefixesInUse.has(numbered)) return numbered;
    }
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
  const titleWords = title
    .toLowerCase()
    .replace(/[^a-z ]/g, "")
    .split(" ")
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
