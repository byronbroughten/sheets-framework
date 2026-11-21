import type { RawSource } from "./RawSource";

let installed: RawSource | undefined;

export function installRawSource(source: RawSource): void {
  installed = source;
}

export function hasInstalledRawSource(): boolean {
  return installed !== undefined;
}

export function installedRawSource(): RawSource {
  if (installed === undefined) {
    throw new Error(
      "RawSource has not been installed. The host must construct GoogleSheetsAPI before SpreadsheetRaw.init.",
    );
  }
  return installed;
}
