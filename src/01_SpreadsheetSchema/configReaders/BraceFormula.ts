import { cellFormula } from "./BraceFormula/cellFormula.ts";
import { formulaNames } from "./BraceFormula/formulaNames.ts";
import { ruleFormula } from "./BraceFormula/ruleFormula.ts";

export type * from "./BraceFormula/cellFormula.ts";
export type * from "./BraceFormula/formulaNames.ts";
export type * from "./BraceFormula/ruleFormula.ts";

export const BraceFormula = {
  ...cellFormula,
  ...ruleFormula,
  ...formulaNames,
};
