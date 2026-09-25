import { ssConfigGet } from "./spreadsheetConfigTypes";

type DimensionKind = "c" | "r";

export const dimensionIds = {
  col(idPrefix: string, suffix?: string): string {
    return dimensionId("c", idPrefix, suffix);
  },
  row(idPrefix: string, suffix?: string): string {
    return dimensionId("r", idPrefix, suffix);
  },
  // Undefined for an id minted under a different delimiter, which parses as one part.
  colIdPrefixOrUndefined(colId: string): string | undefined {
    const parts = colId.split(ssConfigGet("idDelimiter"));
    const idPrefix = parts[1];
    if (parts.length !== 3 || parts[0] !== "c" || !idPrefix || !parts[2]) {
      return undefined;
    }
    return idPrefix;
  },
};

function dimensionId(
  kind: DimensionKind,
  idPrefix: string,
  suffix: string = randomSuffix(),
): string {
  if (!idPrefix) {
    throw new Error(`Attempted to make id for sheet without an idPrefix`);
  }
  return [kind, idPrefix, suffix].join(ssConfigGet("idDelimiter"));
}

function randomSuffix(): string {
  const length = 7;
  const alphabet =
    "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ_-";
  let result = "";
  for (let i = 0; i < length; i++) {
    result += alphabet.charAt(Math.floor(Math.random() * alphabet.length));
  }
  return result;
}
