import type { ColumnConfigsGeneric } from "../01_SpreadsheetSchema/makeConfigs";

function oneLineJsonObject(record: object): string {
  const fields = Object.entries(record).map(
    ([key, value]) => `${JSON.stringify(key)}: ${JSON.stringify(value)}`,
  );
  return `{ ${fields.join(", ")} }`;
}

export function columnConfigsFileSource(
  columnConfigs: ColumnConfigsGeneric,
): string {
  const sheets = Object.entries(columnConfigs);
  if (sheets.length === 0) {
    return "{}";
  }
  const blocks = sheets.map(([sheetName, tableColumnConfigs]) => {
    const columnLines = Object.entries(tableColumnConfigs).map(
      ([columnName, columnConfig]) =>
        `    ${JSON.stringify(columnName)}: ${oneLineJsonObject(columnConfig)}`,
    );
    return `  ${JSON.stringify(sheetName)}: {\n${columnLines.join(",\n")}\n  }`;
  });
  return `{\n${blocks.join(",\n")}\n}`;
}

export function oneLinePerEntryFileSource(
  configs: Record<string, object>,
): string {
  const entries = Object.entries(configs);
  if (entries.length === 0) {
    return "{}";
  }
  const lines = entries.map(
    ([key, config]) => `  ${JSON.stringify(key)}: ${oneLineJsonObject(config)}`,
  );
  return `{\n${lines.join(",\n")}\n}`;
}
