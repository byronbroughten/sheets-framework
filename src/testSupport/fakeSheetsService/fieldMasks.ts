export const fieldMasks = {
  replay(
    fieldMask: string | undefined,
    requestKind: string,
    fieldReplays: Record<string, () => void>,
  ): void {
    (fieldMask ?? "").split(",").forEach((rawField) => {
      const field = rawField.trim();
      const fieldReplay = fieldReplays[field];
      if (fieldReplay === undefined) {
        throw new Error(
          `The fake Sheets service does not replay ${requestKind} field "${field}".`,
        );
      }
      fieldReplay();
    });
  },
};
