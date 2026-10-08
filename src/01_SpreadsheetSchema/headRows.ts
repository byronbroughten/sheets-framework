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
function headRowIndex(role: HeadRole): number {
  return indexAtOffset(tableLayout.headRowOffsets[role]);
}

function indexAtOffset(offset: number): number {
  return -1 - offset;
}

// The largest offset, not a named role, so the extent follows tableLayout.
const countAboveHeader = Math.max(...Object.values(tableLayout.headRowOffsets));

// One per row, so a row two roles share appears once.
const headRowIndexes: ReadonlySet<number> = new Set(
  Obj.keys(tableLayout.headRowOffsets).map(headRowIndex),
);

export const headRows = {
  countAboveHeader,
  topIndex: indexAtOffset(countAboveHeader),
  lastAboveHeaderIndex: indexAtOffset(1),
  index: headRowIndex,
  indexes(): number[] {
    return [...headRowIndexes];
  },
  isIndex(rowIndex: number): boolean {
    return headRowIndexes.has(rowIndex);
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
