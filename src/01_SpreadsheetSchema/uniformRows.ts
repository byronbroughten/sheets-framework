import type { UniformRowName } from "../00_Source/CellValues/cellValues";
import { Obj } from "../utils/Obj";
import { headRows } from "./headRows";

export const uniformRows = {
  indexes(): Record<UniformRowName, number> {
    return {
      columnId: headRows.index("columnId"),
      groupHeading1: headRows.index("groupHeading1"),
      action: headRows.index("action"),
      header: headRows.index("header"),
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
