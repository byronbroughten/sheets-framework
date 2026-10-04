import type { UniformRowName } from "../00_Source/CellValues/cellValues";
import { Obj } from "../utils/Obj";
import { sheetLayout } from "./sheetLayout";

// Table-relative, so a head row sits at a negative index above body row 0.
export const uniformRows = {
  indexes(): Record<UniformRowName, number> {
    const offsets = sheetLayout.headRowOffsets;
    return {
      columnId: headRowIndex(offsets.columnId),
      colGroupName: headRowIndex(offsets.groupHeading1),
      action: headRowIndex(offsets.action),
      tableHeader: headRowIndex(offsets.header),
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
