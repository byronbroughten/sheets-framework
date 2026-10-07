import type {
  CellValue,
  CellValueName,
} from "../00_Source/CellValues/cellValues";
import { Obj } from "../utils/Obj";
import { tableLayout } from "./tableLayout";

type HeadRowOffsets = typeof tableLayout.headRowOffsets;
export type HeadRole = keyof HeadRowOffsets;

interface HeadRoleValueNames extends Record<HeadRole, CellValueName> {
  header: "string";
  action: "boolean";
  groupHeading2: "string";
  groupHeading1: "string";
  columnId: "string";
}

export type HeadRowRoles<HR extends HeadRole> = {
  [RL in HeadRole]: HeadRowOffsets[RL] extends HeadRowOffsets[HR] ? RL : never;
}[HeadRole];

// A row's cells may hold any of its roles' values, so a shared row's type is their union.
export type HeadRowValueName<HR extends HeadRole> =
  HeadRoleValueNames[HeadRowRoles<HR>];
export type HeadRowValue<HR extends HeadRole> = CellValue<HeadRowValueName<HR>>;

// Table-relative, so a head row sits at a negative index above body row 0.
export const headRows = {
  index(role: HeadRole): number {
    return -1 - tableLayout.headRowOffsets[role];
  },
  // One per row, so a row two roles share appears once.
  indexes(): number[] {
    return [...new Set(Obj.keys(tableLayout.headRowOffsets).map(headRows.index))];
  },
  isIndex(rowIndex: number): boolean {
    return headRows.indexes().includes(rowIndex);
  },
  rolesAt(rowIndex: number): [HeadRole, ...HeadRole[]] {
    const [first, ...rest] = Obj.keys(tableLayout.headRowOffsets).filter(
      (role) => headRows.index(role) === rowIndex,
    );
    if (first === undefined) {
      throw new Error(`Row index ${rowIndex} is not a head row.`);
    }
    return [first, ...rest];
  },
  rolesSharing<HR extends HeadRole>(role: HR): HeadRowRoles<HR>[] {
    return Obj.keys(tableLayout.headRowOffsets).filter(
      (other): other is HeadRowRoles<HR> =>
        headRows.index(other) === headRows.index(role),
    );
  },
};
