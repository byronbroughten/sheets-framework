import {
  type HeadRole,
  type HeadRowRoles,
  headRows,
  type HeadRowValue,
} from "../../01_SpreadsheetSchema/headRows";
import { RowCommonRaw } from "./RowCommonRaw";
import type { TableRawProps } from "./TableBaseRaw";

export type HeadRowRawProps<HR extends HeadRole> = TableRawProps & {
  headRole: HR;
};

export class HeadRowBaseRaw<
  HR extends HeadRole = HeadRole,
> extends RowCommonRaw {
  readonly headRole: HR;
  readonly roles: HeadRowRoles<HR>[];
  constructor({ headRole, ...rest }: HeadRowRawProps<HR>) {
    super({ ...rest, rowIndex: headRows.index(headRole) });
    this.headRole = headRole;
    this.roles = headRows.rolesSharing(headRole);
    this.ensureStateExists();
  }
  get workingValueArr(): (HeadRowValue<HR> | "")[] {
    return [...this.rowState.values()].map((cellState) => cellState.value) as (
      HeadRowValue<HR> | ""
    )[];
  }
}
