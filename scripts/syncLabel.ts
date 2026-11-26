// A line of sync's own, set apart from the report and the host's log.
export function labelled(text: string): string {
  return `sync: ${text}`;
}
