import { Val } from "@byronbroughten/utils/val";

import {
  type IntendedTrait,
  isTableDeletion,
} from "../../01_SpreadsheetSchema/configReaders/SyncPlanner";
import { tableConfigsByName } from "../../01_SpreadsheetSchema/configReaders/tableConfigsTypes";
import type { TableRaw } from "../../02_SpreadsheetRaw/TableRaw";
import { SpreadsheetNamed } from "../../04_SpreadsheetNamed/SpreadsheetNamed";
import { SpreadsheetBaseOperator } from "../SpreadsheetBaseOperator";

export class TableConfigsTraitSenderOperator extends SpreadsheetBaseOperator {
  get ss(): SpreadsheetNamed {
    return new SpreadsheetNamed(this.spreadsheetNamedProps);
  }
  // A renamed Table is addressed by its previous key, which the refuser guarantees the previous configs hold.
  updateTableNames(sent: IntendedTrait[]): this {
    for (const intended of sent) {
      if (intended.kind !== "table" || intended.trait !== "name") continue;
      this._previousTable(intended.table).updateTableName(
        String(intended.value),
      );
    }
    return this;
  }
  deleteTables(sent: IntendedTrait[]): this {
    for (const intended of sent) {
      if (intended.kind !== "table" || !isTableDeletion(intended)) continue;
      this._previousTable(intended.table).queueDelete();
    }
    return this;
  }
  // By the previous configs' name, the one an alteration addresses a Table by.
  tableIdOf(tableKey: string): string {
    return Val.assert(
      tableConfigsByName()[tableKey],
      `previous Table ${tableKey}`,
    ).tableId;
  }
  private _previousTable(tableKey: string): TableRaw {
    return this.ss.raw.table(this.tableIdOf(tableKey));
  }
}
