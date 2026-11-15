import { BraceFormula } from "./BraceFormula";
import type {
  SyncNotice,
  SyncOutsideReferencesCheck,
  SyncPlanEffect,
} from "./syncRunResult";

// One API-access Table's formula column, as this run reads it.
export interface FormulaColumn {
  tableKey: string;
  // The column's body rows in A1, on its tab.
  cell: string;
  // In brace form, the text the configs this run writes record.
  formula: string;
}

interface OutsideReferencesProps {
  // A set, so a tab's deletion can hand over every Table on it.
  deletedTables: ReadonlySet<string>;
  formulaColumns: FormulaColumn[];
}

// Live, a deleted Table's references are rewritten to its old body rows, which read blank (#318).
const deletedTableEffect = "blank";

export const OutsideReferences = {
  // One check per deleted Table a formula outside the deleted Tables names; a deleted Table's own formulas go with it.
  checks({
    deletedTables,
    formulaColumns,
  }: OutsideReferencesProps): SyncOutsideReferencesCheck[] {
    const outside = formulaColumns.filter(
      ({ tableKey }) => !deletedTables.has(tableKey),
    );
    return [...deletedTables].flatMap((table) => {
      const references = outside
        .filter((column) =>
          BraceFormula.columnsNamed(column).some(
            (named) => named.tableKey === table,
          ),
        )
        .map(({ cell, formula }) => ({ cell, formula }));
      if (references.length === 0) return [];
      return [
        {
          kind: "outsideReferences",
          subject: { table },
          effect: deletedTableEffect,
          references,
        },
      ];
    });
  },
  effect({ references, effect }: SyncOutsideReferencesCheck): SyncPlanEffect {
    return { kind: "outsideReferences", count: references.length, effect };
  },
  // Given whenever a Table deletion is planned; nothing is read to check the rest.
  uncheckedNotice: {
    kind: "text",
    text: "A Table deletion's outside formulas are checked only in the API-access Tables' formula columns; a formula anywhere else that reads a deleted Table isn't listed.",
  } satisfies SyncNotice,
};
