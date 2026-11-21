import { unreplayed } from "./unreplayed.ts";

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
        throw unreplayed(`${requestKind} field "${field}"`);
      }
      fieldReplay();
    });
  },
};
