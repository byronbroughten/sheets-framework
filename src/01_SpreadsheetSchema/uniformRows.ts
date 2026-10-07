import type { UniformRowName } from "../00_Source/CellValues/cellValues";
import { Obj } from "../utils/Obj";
import { tableLayout } from "./tableLayout";

// Table-relative, so a head row sits at a negative index above body row 0.
export const uniformRows = {
  indexes(): Record<UniformRowName, number> {
    const offsets = tableLayout.headRowOffsets;
    return {
      columnId: headRowIndex(offsets.columnId),
      groupHeading1: headRowIndex(offsets.groupHeading1),
      action: headRowIndex(offsets.action),
      header: headRowIndex(offsets.header),
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

function headRowIndex(offsetAboveHeader: number): number {
  return -1 - offsetAboveHeader;
}
