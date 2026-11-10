import { Str } from "@byronbroughten/utils/str";
import { Val } from "@byronbroughten/utils/val";

export interface IdPrefixLabel {
  label: string;
  idPrefix: string;
}

export class IdPrefixesByTableId {
  constructor(private readonly byTableId: ReadonlyMap<string, string>) {}
  assertIdPrefix(tableId: string): string {
    return Val.assert(this.byTableId.get(tableId), "ID prefix");
  }
  tableIds(): string[] {
    return [...this.byTableId.keys()];
  }
  idPrefixes(): string[] {
    return [...this.byTableId.values()];
  }
  forEach(callback: (idPrefix: string, tableId: string) => void): void {
    this.byTableId.forEach(callback);
  }
}

export interface IdPrefixFromTitleProps {
  title: string;
  prefixesInUse: ReadonlySet<string>;
  retiredPrefixes?: Iterable<string>;
}

export interface RetiredOrGoneProps {
  previouslyRetired: readonly string[];
  previousByTableId: IdPrefixesByTableId;
  liveTableIds: ReadonlySet<string>;
}

export interface RetiredIdPrefixesProps {
  previouslyRetired: readonly string[];
  previousByTableId: IdPrefixesByTableId;
  currentByTableId: IdPrefixesByTableId;
}

const vowels = new Set(["a", "e", "i", "o", "u"]);

export const idPrefixes = {
  // A retired prefix is taken for good: rows elsewhere may still carry IDs minted under it.
  fromTitle({
    title,
    prefixesInUse,
    retiredPrefixes = [],
  }: IdPrefixFromTitleProps): string {
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
    const current = new Set(currentByTableId.idPrefixes());
    const retired = idPrefixes.retiredOrGone({
      previouslyRetired,
      previousByTableId,
      liveTableIds: new Set(currentByTableId.tableIds()),
    });
    return [...retired].filter((idPrefix) => !current.has(idPrefix)).sort();
  },
  // The previous retired list plus the prefix of each previous Table not live now.
  retiredOrGone({
    previouslyRetired,
    previousByTableId,
    liveTableIds,
  }: RetiredOrGoneProps): Set<string> {
    const retired = new Set(previouslyRetired);
    previousByTableId.forEach((idPrefix, tableId) => {
      if (!liveTableIds.has(tableId)) retired.add(idPrefix);
    });
    return retired;
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
