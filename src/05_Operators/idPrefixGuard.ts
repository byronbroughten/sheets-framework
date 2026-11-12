export interface IdPrefixGuardProps {
  previousByTableId: ReadonlyMap<string, string>;
  nextByTableId: ReadonlyMap<string, string>;
  retiredPrefixes: ReadonlySet<string>;
}

// Every column ID and every generated key built on a prefix moves with it, so a change stops the run.
export function assertIdPrefixesUnchanged({
  previousByTableId,
  nextByTableId,
  retiredPrefixes,
}: IdPrefixGuardProps): void {
  const changes = [...nextByTableId].flatMap(([tableId, next]) => {
    const previous = previousByTableId.get(tableId);
    if (previous === undefined) {
      if (!retiredPrefixes.has(next)) return [];
      return [
        `${tableLabel(tableId)} has ID prefix "${next}", which an earlier Table retired.`,
      ];
    }
    if (next === previous) return [];
    return [
      `${tableLabel(tableId)} had ID prefix "${previous}" and is now "${next}".`,
    ];
  });
  if (changes.length === 0) return;
  throw new Error(
    `The regen would change an ID prefix, so it stopped before writing. ${changes.join(" ")}`,
  );
}

function tableLabel(tableId: string): string {
  return `Table "${tableId}"`;
}
