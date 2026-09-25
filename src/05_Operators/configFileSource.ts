import type {
  ColumnConfigsGeneric,
  SheetConfigsBase,
} from "../01_SpreadsheetSchema/makeConfigs";

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

export function sheetConfigsFileSource(sheetConfigs: SheetConfigsBase): string {
  const sheets = Object.entries(sheetConfigs);
  if (sheets.length === 0) {
    return "{}";
  }
  const lines = sheets.map(
    ([sheetName, sheetConfig]) =>
      `  ${JSON.stringify(sheetName)}: ${oneLineJsonObject(sheetConfig)}`,
  );
  return `{\n${lines.join(",\n")}\n}`;
}
