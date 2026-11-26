import { unreplayed } from "../fakeSheetsService/unreplayed.ts";

interface KeyCheck {
  keys: string[];
  modelled: string[];
  kind: "query parameter" | "body key";
}

export function validateKeys({ keys, modelled, kind }: KeyCheck): void {
  const other = keys.find((key) => !modelled.includes(key));
  if (other !== undefined) throw unreplayed(`the ${kind} "${other}"`);
}
