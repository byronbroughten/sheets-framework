import type { Chore } from "./Chore";

export const fillMissingRowIds: Chore = {
  description:
    "Writes a generated row ID into every blank ID cell, on every sheet the config gives an ID prefix.",
  action: (ss) => {
    ss.fillMissingRowIds();
    ss.batchUpdateGSheets();
  },
};
