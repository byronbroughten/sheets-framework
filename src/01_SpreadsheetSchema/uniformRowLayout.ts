import {
  spreadsheetConfigColumnLabel,
  spreadsheetConfigIndexHeaders,
} from "./spreadsheetConfigFields";

export interface UniformRowLayoutIndexes {
  columnIdRowIdxBase0: number;
  columnGroupHeadingRowIndexBase0: number;
  actionRowIndexBase0: number;
  tableHeaderRowIndexBase0: number;
}

export type UniformRowLayoutKey = keyof UniformRowLayoutIndexes;

const uniformRowLayoutKeys: readonly UniformRowLayoutKey[] = [
  "columnIdRowIdxBase0",
  "columnGroupHeadingRowIndexBase0",
  "actionRowIndexBase0",
  "tableHeaderRowIndexBase0",
];

const labels = uniformRowLayoutColumnLabels();

export const uniformRowLayout = {
  labels,
  validate(
    config: UniformRowLayoutIndexes,
    rowLabels: Record<UniformRowLayoutKey, string> = labels,
  ): void {
    const firstDataRowIndex = config.tableHeaderRowIndexBase0 + 1;
    const nameByIndex = new Map<number, string>();
    uniformRowLayoutKeys.forEach((key) => {
      const index = config[key];
      const label = rowLabels[key];
      if (typeof index !== "number" || !Number.isInteger(index) || index < 0) {
        throw new Error(
          `${label} must be an integer ≥ 0, got ${JSON.stringify(index)}.`,
        );
      }
      const existing = nameByIndex.get(index);
      if (existing !== undefined) {
        throw new Error(`${existing} and ${label} must not share a row.`);
      }
      if (index === firstDataRowIndex) {
        throw new Error(`${label} must not land on the first data row.`);
      }
      nameByIndex.set(index, label);
    });
  },
};

function uniformRowLayoutColumnLabels(): Record<UniformRowLayoutKey, string> {
  return uniformRowLayoutKeys.reduce(
    (labels, key) => {
      labels[key] = spreadsheetConfigColumnLabel(
        spreadsheetConfigIndexHeaders[key],
      );
      return labels;
    },
    {} as Record<UniformRowLayoutKey, string>,
  );
}
