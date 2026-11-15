export function unreplayed(what: string): Error {
  return new Error(`The fake Sheets service does not replay ${what}.`);
}
