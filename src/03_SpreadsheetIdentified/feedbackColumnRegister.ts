// Column ids by tableId; installed as data, since the endpoints that declare them sit tiers above.
export type FeedbackColumnIds = ReadonlyMap<string, ReadonlySet<string>>;

let installed: FeedbackColumnIds | undefined;

// Mirrors installConfigs, so two spreadsheets in one program can't disagree about a blank row.
export function installFeedbackColumnIds(ids: FeedbackColumnIds): void {
  if (installed !== undefined && !isSameFeedbackColumnIds(installed, ids)) {
    throw new Error(
      "A different set of feedback columns is already installed. A program installs one set of endpoints, once.",
    );
  }
  installed = ids;
}

export function installedFeedbackColumnIds(): FeedbackColumnIds {
  return installed ?? new Map();
}

function isSameFeedbackColumnIds(
  a: FeedbackColumnIds,
  b: FeedbackColumnIds,
): boolean {
  if (a.size !== b.size) return false;
  return [...a].every(([tableId, columnIds]) => {
    const other = b.get(tableId);
    return (
      other !== undefined &&
      other.size === columnIds.size &&
      [...columnIds].every((columnId) => other.has(columnId))
    );
  });
}
