import type { UniformRowName } from "../00_Source/CellValues/cellValues";
import { Obj } from "../utils/Obj";
import { sheetLayout } from "./sheetLayout";

export const uniformRows = {
  indexes(): Record<UniformRowName, number> {
    return {
      columnId: sheetLayout.colIdRowIndex,
      colGroupName: sheetLayout.colGroupHeadingRowIndex,
      action: sheetLayout.actionRowIndex,
      tableHeader: sheetLayout.tableHeaderRowIndex,
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
