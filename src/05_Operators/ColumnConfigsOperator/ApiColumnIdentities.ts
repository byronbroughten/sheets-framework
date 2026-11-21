import { Val } from "@byronbroughten/utils/val";

import { columnNameOfHeader } from "../../01_SpreadsheetSchema/configReaders/columnConfigsTypes";
import type { ColumnProfileRaw } from "../../02_SpreadsheetRaw/ColumnProfileRaw";
import type { SpreadsheetNamed } from "../../04_SpreadsheetNamed/SpreadsheetNamed";
import { PreviousColumnConfigs } from "../PreviousColumnConfigs";
import type { TableConfigsOperator } from "../TableConfigsOperator";

export interface ColumnIdentity {
  tableId: string;
  columnId: string;
}

export interface NamedColumnIdentity extends ColumnIdentity {
  tableKey: string;
  columnName: string;
  previousEmptyValueAllowed: boolean;
}

interface ApiColumnIdentitiesProps {
  ss: SpreadsheetNamed;
  tableConfigOperator: TableConfigsOperator;
}

// The columns of every API-access Table, by ID and by previous name, which each ColumnConfigsOperator collaborator reads.
export class ApiColumnIdentities {
  private readonly ss: SpreadsheetNamed;
  private readonly tableConfigOperator: TableConfigsOperator;
  constructor({ ss, tableConfigOperator }: ApiColumnIdentitiesProps) {
    this.ss = ss;
    this.tableConfigOperator = tableConfigOperator;
  }
  // Derived fresh each call, not cached — a stored field goes stale across
  // this coordinator's per-access getter rebuilds.
  get tableIdsApiAccesses(): Set<string> {
    return new Set(this.tableConfigOperator.tableIdsApiAccesses());
  }
  // Per Table, the previous configs' order first, so a quiet run leaves the generated files unchanged.
  columnIdentities(): ColumnIdentity[] {
    return [...this.tableIdsApiAccesses].flatMap((tableId) => {
      const { profile } = this.ss.raw.table(tableId);
      const liveColumnIds = new Set(profile.columnIds);
      const previousColumnIds = new PreviousColumnConfigs(
        tableId,
      ).columnIds.filter((columnId) => liveColumnIds.has(columnId));
      const previous = new Set(previousColumnIds);
      const newColumnIds = profile.columnIds.filter(
        (columnId) => !previous.has(columnId),
      );
      return [...previousColumnIds, ...newColumnIds].map((columnId) => ({
        tableId,
        columnId,
      }));
    });
  }
  // Keyed by previous names through the IDs, so a rename keeps an alteration's target (ADR 0035).
  namedColumnIdentities(): NamedColumnIdentity[] {
    const tableKeysByTableId = this.tableConfigOperator.tableKeysByTableId();
    return this.columnIdentities().map((identity) => {
      const previous = new PreviousColumnConfigs(identity.tableId);
      return {
        ...identity,
        tableKey: Val.assert(
          previous.tableKey ?? tableKeysByTableId.get(identity.tableId),
          "API-access Table key",
        ),
        columnName:
          previous.columnName(identity.columnId) ??
          columnNameOfHeader(this.describedColumn(identity).header),
        previousEmptyValueAllowed: previous.emptyValueAllowed(
          identity.columnId,
        ),
      };
    });
  }
  describedColumn({ tableId, columnId }: ColumnIdentity): ColumnProfileRaw {
    return this.ss.raw.table(tableId).profile.columnById(columnId);
  }
}
