// PROTOTYPE for sheets-framework#49: throwaway and type-only. It never reaches master.
// The formula-rewrite chore, written against #48's surviving candidate (ii): Table primary, through `ss.table(name)`.
// Its targets are real: sheets-real-estate formula columns that still use same-row A1 references,
// read from each Table's first body row on 2026-10-01 (404 of 430 formulas there are already Table notation).
// Compared: (A) template strings from column members, (A') literal text, (B) a `formula` tagged template.
import type { ActionReturn } from "../PROTOTYPE_issue48_tableCandidates/shared";

// A trimmed copy of the app's schema: config key, header text, and whether gen:configs sampled a formula.
interface Tables {
  household: {
    id: { header: "ID"; isFormula: false };
    residentFullNames: { header: "Resident full names"; isFormula: true };
    nextRentChangeDate: { header: "Next rent change date"; isFormula: true };
    nextMinusLastIncreaseYears: {
      header: "Next minus last increase years";
      isFormula: true;
    };
    increaseBy3PerYear: { header: "Increase by 3% per year"; isFormula: true };
    daysInNextRentIncreaseYear: {
      header: "Days in next rent increase year";
      isFormula: true;
    };
  };
  resident: {
    fullName: { header: "Full name"; isFormula: true };
    householdId: { header: "Household ID"; isFormula: false };
    nameRank: { header: "Name rank"; isFormula: false };
  };
  pet: {
    birthDateApproximate: {
      header: "Birth date, approximate";
      isFormula: false;
    };
    age: { header: "Age"; isFormula: true };
  };
  furnace: {
    name: { header: "Name"; isFormula: true };
    propertyId: { header: "Property ID"; isFormula: false };
    propertyName: { header: "Property name"; isFormula: true };
    descriptor: { header: "Descriptor"; isFormula: false };
  };
  property: {
    id: { header: "ID"; isFormula: false };
    name: { header: "Name"; isFormula: false };
  };
  variable: {
    today: { header: "Today"; isFormula: true };
  };
}
type TableName = keyof Tables;
type ColName<TN extends TableName> = keyof Tables[TN] & string;
type IsFormula<
  TN extends TableName,
  CN extends ColName<TN>,
> = Tables[TN][CN] extends { isFormula: true } ? true : false;

interface SpreadsheetNamed {
  table<TN extends TableName>(tableName: TN): TableNamed<TN>;
}

interface TableNamed<TN extends TableName> {
  // The live Table name: a reference needs it, and it drifts from the config key (`addOccChargeOnetime`).
  readonly name: string;
  column<CN extends ColName<TN>>(col: CN): TableColumnNamed<TN, CN>;
}

// What any Table column offers a formula, whatever its Table.
interface FormulaSource {
  // `household[ID]`: the whole body column; also what a dropdown validator source would be.
  readonly reference: string;
  // `SINGLE(household[ID])`: this row's value, on whichever body row the formula sits.
  readonly single: string;
}

interface TableColumnNamed<TN extends TableName, CN extends ColName<TN>>
  extends FormulaSource {
  readonly headerText: Tables[TN][CN]["header" & keyof Tables[TN][CN]];
  readonly meta: TableColumnMetaNamed<TN, CN>;
  // One text for every body row, as `pasteData` already sends it (byro-repo#38); `updateCells` breaks a Table reference.
  updateAllFormulas(
    formula: IsFormula<TN, CN> extends true ? string | Formula : never,
  ): this;
}

// The top body row's formula text, so a dry run can print before and after; today only `isFormula` is read.
interface TableColumnMetaNamed<TN extends TableName, CN extends ColName<TN>> {
  readonly activeFormula: IsFormula<TN, CN> extends true ? string : never;
}

// (B) only.
type Formula = string & { readonly __formula: true };
interface Single {
  readonly __single: true;
}
type FormulaPart = FormulaSource | Single | string | number;
declare function formula(
  strings: TemplateStringsArray,
  ...parts: FormulaPart[]
): Formula;
declare function single(column: FormulaSource): Single;

interface Chore {
  description: string;
  action: (ss: SpreadsheetNamed) => ActionReturn | void;
}

const description =
  "Rewrite the formula columns that still use same-row A1 references in Table notation.";

// (A) Template strings from column members; each comment is the column's A1 formula today.
export const tableNotationFormulasA: Chore = {
  description,
  action(ss) {
    const household = ss.table("household");
    const resident = ss.table("resident");
    const furnace = ss.table("furnace");
    const property = ss.table("property");
    const pet = ss.table("pet");
    const variable = ss.table("variable");

    // =$AG5 * 0.03
    household
      .column("increaseBy3PerYear")
      .updateAllFormulas(
        `=${household.column("nextMinusLastIncreaseYears").single} * 0.03`,
      );

    // =DATE(YEAR($R5), 12, 31) - DATE(YEAR($R5), 1, 1) + 1
    const nextChange = household.column("nextRentChangeDate").single;
    household
      .column("daysInNextRentIncreaseYear")
      .updateAllFormulas(
        `=DATE(YEAR(${nextChange}), 12, 31) - DATE(YEAR(${nextChange}), 1, 1) + 1`,
      );

    // =TEXTJOIN(", ",TRUE, SORT(FILTER(resident[Full name], resident[Household ID]=$B5), FILTER(resident[Name rank], …), TRUE))
    const ofThisHousehold = `${resident.column("householdId").reference}=${household.column("id").single}`;
    household
      .column("residentFullNames")
      .updateAllFormulas(
        `=TEXTJOIN(", ", TRUE, SORT(` +
          `FILTER(${resident.column("fullName").reference}, ${ofThisHousehold}), ` +
          `FILTER(${resident.column("nameRank").reference}, ${ofThisHousehold}), TRUE))`,
      );

    // =FLOOR((variable[Today] - G5)/365)
    pet
      .column("age")
      .updateAllFormulas(
        `=FLOOR((${variable.column("today").reference} - ${pet.column("birthDateApproximate").single})/365)`,
      );

    // =CONCATENATE($D5, ", ", $F5)
    furnace
      .column("name")
      .updateAllFormulas(
        `=CONCATENATE(${furnace.column("propertyName").single}, ", ", ${furnace.column("descriptor").single})`,
      );

    // =FILTER(property[Name], property[ID]=$C5)
    furnace
      .column("propertyName")
      .updateAllFormulas(
        `=FILTER(${property.column("name").reference}, ${property.column("id").reference}=${furnace.column("propertyId").single})`,
      );
  },
};

// (A') Literal text: the author types the Table notation the sheet already uses 404 times.
export const tableNotationFormulasLiteral: Chore = {
  description,
  action(ss) {
    ss.table("household")
      .column("increaseBy3PerYear")
      .updateAllFormulas(
        "=SINGLE(household[Next minus last increase years]) * 0.03",
      );
    // A typo such as `property[Nmae]` passes tsc and lands in the sheet as an error.
    ss.table("furnace")
      .column("propertyName")
      .updateAllFormulas(
        "=FILTER(property[Name], property[ID]=SINGLE(furnace[Property ID]))",
      );
  },
};

// (B) A `formula` tag: a column interpolates as its reference, `single(column)` as its row value.
export const tableNotationFormulasB: Chore = {
  description,
  action(ss) {
    const household = ss.table("household");
    const furnace = ss.table("furnace");
    const property = ss.table("property");
    household
      .column("increaseBy3PerYear")
      .updateAllFormulas(
        formula`=${single(household.column("nextMinusLastIncreaseYears"))} * 0.03`,
      );
    furnace
      .column("propertyName")
      .updateAllFormulas(
        formula`=FILTER(${property.column("name")}, ${property.column("id")}=${single(furnace.column("propertyId"))})`,
      );
  },
};

export function dryRunLine(ss: SpreadsheetNamed, after: string): string {
  const column = ss.table("household").column("increaseBy3PerYear");
  return `${column.headerText}: ${column.meta.activeFormula} -> ${after}`;
}

export function refusedOnDataColumn(ss: SpreadsheetNamed): void {
  // @ts-expect-error -- `id` is not a formula column
  ss.table("household").column("id").updateAllFormulas("=1");
}
