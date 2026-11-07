type ExtendedValue = GoogleAppsScript.Sheets.Schema.ExtendedValue;

const comparedKinds = [
  "stringValue",
  "numberValue",
  "boolValue",
  "formulaValue",
] as const;

export const extendedValue = {
  label(value: ExtendedValue | undefined): string {
    if (!value) return "(no value)";
    const { stringValue, numberValue, boolValue, formulaValue } = value;
    if (stringValue !== undefined) return `"${stringValue}"`;
    if (numberValue !== undefined) return String(numberValue);
    if (boolValue !== undefined) return String(boolValue);
    if (formulaValue !== undefined) return formulaValue;
    return "(no value)";
  },
  // Plain text, unquoted, for a value that names a row rather than shows a change.
  text(value: ExtendedValue | undefined): string | undefined {
    const shown =
      value?.stringValue ?? value?.numberValue ?? value?.boolValue ?? undefined;
    return shown === undefined ? undefined : String(shown);
  },
  isEqual(a: ExtendedValue, b: ExtendedValue): boolean {
    return comparedKinds.some(
      (kind) => a[kind] !== undefined && a[kind] === b[kind],
    );
  },
};
