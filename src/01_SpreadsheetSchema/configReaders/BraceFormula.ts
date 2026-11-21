import { cellFormula } from "./BraceFormula/cellFormula";
import { formulaNames } from "./BraceFormula/formulaNames";
import { ruleFormula } from "./BraceFormula/ruleFormula";

export type {
  BraceConfigs,
  BraceNames,
  ColumnTypeProps,
} from "./BraceFormula/cellFormula";
export type {
  FormulaRenames,
  NamedColumn,
  RenamedTable,
} from "./BraceFormula/formulaNames";
export type {
  RuleDecompileProps,
  RuleLayout,
} from "./BraceFormula/ruleFormula";

export const BraceFormula = {
  ...cellFormula,
  ...ruleFormula,
  ...formulaNames,
};
