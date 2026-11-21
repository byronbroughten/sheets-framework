import { cellFormula } from "./BraceFormula/cellFormula";
import { formulaNames } from "./BraceFormula/formulaNames";
import { ruleFormula } from "./BraceFormula/ruleFormula";

export type * from "./BraceFormula/cellFormula";
export type * from "./BraceFormula/formulaNames";
export type * from "./BraceFormula/ruleFormula";

export const BraceFormula = {
  ...cellFormula,
  ...ruleFormula,
  ...formulaNames,
};
