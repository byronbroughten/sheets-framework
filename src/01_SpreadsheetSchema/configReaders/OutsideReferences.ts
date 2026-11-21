import { BraceFormula } from "./BraceFormula.ts";
import type {
  SyncNotice,
  SyncOutsideReferenceEffect,
  SyncOutsideReferencesCheck,
  SyncPlanEffect,
  SyncSubject,
} from "./syncRunResult.ts";

// One API-access Table's formula column, as this run reads it.
export interface FormulaColumn {
  tableKey: string;
  // The column's body rows in A1, on its tab.
  cell: string;
  // In brace form, the text the configs this run writes record.
  formula: string;
}

export interface TabTable {
  tableKey: string;
  // The name Sheets' `Table[Header]` references use.
  tableName: string;
}

export interface DeletedTab {
  tab: string;
  title: string;
  tables: TabTable[];
}

interface OutsideReferencesProps {
  deletedTables: ReadonlySet<string>;
  deletedTabs: DeletedTab[];
  formulaColumns: FormulaColumn[];
}

interface CheckProps {
  subject: SyncSubject;
  effect: SyncOutsideReferenceEffect;
  formulaColumns: FormulaColumn[];
  isReference: (column: FormulaColumn) => boolean;
}

const deletionEffects = {
  // Live, a deleted Table's references are rewritten to its old body rows, which read blank (#318).
  table: "blank",
  tab: "ref",
} as const satisfies Record<string, SyncOutsideReferenceEffect>;

export const OutsideReferences = {
  // One check per deleted Table or tab a formula outside every deleted Table names; a deleted Table's own formulas go with it.
  checks({
    deletedTables,
    deletedTabs,
    formulaColumns,
  }: OutsideReferencesProps): SyncOutsideReferencesCheck[] {
    const goneTables = new Set([
      ...deletedTables,
      ...deletedTabs.flatMap(({ tables }) =>
        tables.map(({ tableKey }) => tableKey),
      ),
    ]);
    const outside = formulaColumns.filter(
      ({ tableKey }) => !goneTables.has(tableKey),
    );
    const tableChecks = [...deletedTables].flatMap((table) =>
      checkOf({
        subject: { table },
        effect: deletionEffects.table,
        formulaColumns: outside,
        isReference: (column) => namesTables(column, new Set([table])),
      }),
    );
    const tabChecks = deletedTabs.flatMap((deletedTab) =>
      checkOf({
        subject: { tab: deletedTab.tab },
        effect: deletionEffects.tab,
        formulaColumns: outside,
        isReference: (column) => namesTab(column, deletedTab),
      }),
    );
    return [...tableChecks, ...tabChecks];
  },
  effect({ references, effect }: SyncOutsideReferencesCheck): SyncPlanEffect {
    return { kind: "outsideReferences", count: references.length, effect };
  },
  // Given whenever a Table or tab deletion is planned; nothing is read to check the rest.
  uncheckedNotice: {
    kind: "text",
    text: "A deletion's outside formulas are checked only in the API-access Tables' formula columns; a formula anywhere else that reads a deleted Table or tab isn't listed.",
  } satisfies SyncNotice,
};

function checkOf({
  subject,
  effect,
  formulaColumns,
  isReference,
}: CheckProps): SyncOutsideReferencesCheck[] {
  const references = formulaColumns
    .filter(isReference)
    .map(({ cell, formula }) => ({ cell, formula }));
  if (references.length === 0) return [];
  return [{ kind: "outsideReferences", subject, effect, references }];
}

function namesTables(
  column: FormulaColumn,
  tableKeys: ReadonlySet<string>,
): boolean {
  return BraceFormula.columnsNamed(column).some(({ tableKey }) =>
    tableKeys.has(tableKey),
  );
}

// Sheets matches a sheet reference's title regardless of case.
function namesTab(
  column: FormulaColumn,
  { title, tables }: DeletedTab,
): boolean {
  if (namesTables(column, new Set(tables.map(({ tableKey }) => tableKey)))) {
    return true;
  }
  const tableNames = new Set(tables.map(({ tableName }) => tableName));
  if (
    BraceFormula.liveTablesNamed(column).some((name) => tableNames.has(name))
  ) {
    return true;
  }
  const lowerTitle = title.toLowerCase();
  return BraceFormula.sheetsNamed(column).some(
    (named) => named.toLowerCase() === lowerTitle,
  );
}
