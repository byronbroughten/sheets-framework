import type { UniformRowName } from "../00_Source/CellValues/cellValues";
import { Obj } from "../utils/Obj";
import { ssConfigGet } from "./spreadsheetConfigTypes";

export const uniformRows = {
  indexes(): Record<UniformRowName, number> {
    return {
      columnId: ssConfigGet("columnIdRowIdxBase0"),
      colGroupName: ssConfigGet("columnGroupHeadingRowIndexBase0"),
      action: ssConfigGet("actionRowIndexBase0"),
      tableHeader: ssConfigGet("tableHeaderRowIndexBase0"),
    };
  },
  index(name: UniformRowName): number {
    return uniformRows.indexes()[name];
  },
  nameByIndex(): Map<number, UniformRowName> {
    const indexes = uniformRows.indexes();
    return new Map(
      Obj.keys(indexes).map((name) => [indexes[name], name]),
    ) as Map<number, UniformRowName>;
  },
};
