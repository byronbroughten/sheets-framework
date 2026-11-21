import { Val } from "@byronbroughten/utils/val";

import type { IntendedTrait } from "../../01_SpreadsheetSchema/configReaders/SyncPlanner";
import type { SpreadsheetNamed } from "../../04_SpreadsheetNamed/SpreadsheetNamed";
import type {
  ApiColumnIdentities,
  ColumnIdentity,
  NamedColumnIdentity,
} from "./ApiColumnIdentities";

export interface SentColumnTrait extends NamedColumnIdentity {
  value: IntendedTrait["value"];
}

interface LiveColumnFinderProps {
  ss: SpreadsheetNamed;
  apiColumnIdentities: ApiColumnIdentities;
}

// The live columns a run's sent traits and `before`s name, which ColumnConfigsOperator and its layout operator both look up.
export class LiveColumnFinder {
  private readonly ss: SpreadsheetNamed;
  private readonly apiColumnIdentities: ApiColumnIdentities;
  constructor({ ss, apiColumnIdentities }: LiveColumnFinderProps) {
    this.ss = ss;
    this.apiColumnIdentities = apiColumnIdentities;
  }
  sentColumnTraits(
    sent: IntendedTrait[],
    trait: "header" | "columnType" | "formula" | "deleted",
  ): SentColumnTrait[] {
    const identities = this.apiColumnIdentities.namedColumnIdentities();
    return sent.flatMap((intended) => {
      if (intended.kind !== "column" || intended.trait !== trait) return [];
      const { table: tableKey, column, value } = intended;
      const identity = Val.assert(
        identities.find(
          (named) => named.tableKey === tableKey && named.columnName === column,
        ),
        `live column ${tableKey}.${column}`,
      );
      return [{ ...identity, value }];
    });
  }
  columnIdOf(tableKey: string, columnName: string): string {
    return Val.assert(
      this.columnNamed(tableKey, columnName),
      `live column ${tableKey}.${columnName}`,
    ).columnId;
  }
  // By previous names, as a `before` gives them.
  columnNamed(
    tableKey: string,
    columnName: string,
  ): NamedColumnIdentity | undefined {
    return this.apiColumnIdentities
      .namedColumnIdentities()
      .find(
        (identity) =>
          identity.tableKey === tableKey && identity.columnName === columnName,
      );
  }
  colIndexOf(tableId: string, columnId: string): number {
    return this.ss.raw.table(tableId).columnResolver.colIndexOf(columnId);
  }
  isFirstColumn({ tableId, columnId }: ColumnIdentity): boolean {
    return this.colIndexOf(tableId, columnId) === 0;
  }
}
