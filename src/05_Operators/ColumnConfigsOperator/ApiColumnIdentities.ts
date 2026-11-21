import { Val } from "@byronbroughten/utils/val";

import { columnNameOfHeader } from "../../01_SpreadsheetSchema/configReaders/columnConfigsTypes.ts";
import type { IntendedTrait } from "../../01_SpreadsheetSchema/configReaders/SyncPlanner.ts";
import type { ColumnProfileRaw } from "../../02_SpreadsheetRaw/ColumnProfileRaw.ts";
import type { SpreadsheetNamed } from "../../04_SpreadsheetNamed/SpreadsheetNamed.ts";
import { PreviousColumnConfigs } from "../PreviousColumnConfigs.ts";
import type { TableConfigsOperator } from "../TableConfigsOperator.ts";

export interface ColumnIdentity {
  tableId: string;
  columnId: string;
}

export interface NamedColumnIdentity extends ColumnIdentity {
  tableKey: string;
  columnName: string;
  previousEmptyValueAllowed: boolean;
}

export interface SentColumnTrait extends NamedColumnIdentity {
  value: IntendedTrait["value"];
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
  sentColumnTraits(
    sent: IntendedTrait[],
    trait: "header" | "columnType" | "formula" | "deleted",
  ): SentColumnTrait[] {
    const identities = this.namedColumnIdentities();
    return sent.flatMap((intended) => {
      if (intended.kind !== "column" || intended.trait !== trait) return [];
      const { table: tableKey, column, value } = intended;
      const identity = Val.assert(
        identityNamed(identities, tableKey, column),
        liveColumnLabel(tableKey, column),
      );
      return [{ ...identity, value }];
    });
  }
  columnIdOf(tableKey: string, columnName: string): string {
    return Val.assert(
      this.identityNamed(tableKey, columnName),
      liveColumnLabel(tableKey, columnName),
    ).columnId;
  }
  // By previous names, as a `before` gives them.
  identityNamed(
    tableKey: string,
    columnName: string,
  ): NamedColumnIdentity | undefined {
    return identityNamed(this.namedColumnIdentities(), tableKey, columnName);
  }
  colIndexOf(tableId: string, columnId: string): number {
    return this.ss.raw.table(tableId).columnResolver.colIndexOf(columnId);
  }
  isFirstColumn({ tableId, columnId }: ColumnIdentity): boolean {
    return this.colIndexOf(tableId, columnId) === 0;
  }
}

function identityNamed(
  identities: NamedColumnIdentity[],
  tableKey: string,
  columnName: string,
): NamedColumnIdentity | undefined {
  return identities.find(
    (identity) =>
      identity.tableKey === tableKey && identity.columnName === columnName,
  );
}

function liveColumnLabel(tableKey: string, columnName: string): string {
  return `live column ${tableKey}.${columnName}`;
}
